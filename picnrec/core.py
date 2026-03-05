"""Device interface, image decoding, and export for PicNRec AIO.

Interfaces with the PicNRec's CH340 chip via USB.
If you're having intermittent read issues, try a USB-A to USB-C cable instead of C-to-C
"""

import subprocess
import time
from itertools import batched
from pathlib import Path
from typing import Self

import serial
import serial.tools.list_ports
from PIL import Image

type Palette = list[tuple[int, int, int]]

BAUDRATE = 1_000_000
TIMEOUT = 2.0

IMAGE_WIDTH = 128
IMAGE_HEIGHT = 112
TILES_X = 16   # 128 / 8
TILES_Y = 14   # 112 / 8
TILE_SIZE = 16  # bytes per 8x8 tile (2bpp)
IMAGE_DATA_SIZE = TILES_X * TILES_Y * TILE_SIZE  # 3584
MIN_VALID_BYTES = 50  # non-0xFF bytes required to consider a slot filled

BITMAP_SIZE = 2340
MAX_IMAGES = 18720

# DMG (original Game Boy) green
PALETTE_DMG: Palette = [
    (155, 188, 15),   # lightest
    (139, 172, 15),
    (48, 98, 48),
    (15, 56, 15),     # darkest
]

# Game Boy Pocket: neutral grayscale with warm LCD tint
PALETTE_POCKET: Palette = [
    (174, 166, 145),
    (136, 123, 106),
    (96, 84, 68),
    (78, 63, 42),
]

PALETTE_GRAYSCALE: Palette = [
    (255, 255, 255),
    (170, 170, 170),
    (85, 85, 85),
    (0, 0, 0),
]

PALETTE_INVERTED: Palette = [
    (0, 0, 0),
    (85, 85, 85),
    (170, 170, 170),
    (255, 255, 255),
]

# High contrast: no grays
PALETTE_HARSH: Palette = [
    (255, 255, 255),
    (255, 255, 255),
    (0, 0, 0),
    (0, 0, 0),
]

PALETTES: dict[str, Palette] = {
    'dmg': PALETTE_DMG,
    'pocket': PALETTE_POCKET,
    'grayscale': PALETTE_GRAYSCALE,
    'inverted': PALETTE_INVERTED,
    'harsh': PALETTE_HARSH,
}


class PicNRecDevice:
    """Interface for PicNRec AIO cartridge via direct USB connection.

    Uses the ASCII protocol: !, A{addr}\\0, R, 1, 0.
    """

    __slots__ = ('port', 'serial', 'connected')

    def __init__(self, port: str | None = None) -> None:
        self.port = port
        self.serial: serial.Serial | None = None
        self.connected = False

    def __repr__(self) -> str:
        return f"PicNRecDevice(port={self.port!r}, connected={self.connected})"

    def __enter__(self) -> Self:
        self.connect()
        return self

    def __exit__(self, *exc: object) -> None:
        self.disconnect()

    def find_device(self) -> str | None:
        return next(
            (p.device for p in serial.tools.list_ports.comports()
             if p.vid == 0x1A86 and p.pid == 0x7523),
            None,
        )

    def connect(self, port: str | None = None) -> None:
        if port:
            self.port = port
        elif not self.port:
            self.port = self.find_device()

        if not self.port:
            raise ConnectionError("No PicNRec device found. Check USB connection.")

        try:
            self.serial = serial.Serial(
                port=self.port,
                baudrate=BAUDRATE,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=TIMEOUT,
            )
            self.serial.dtr = True
            self.serial.rts = True
            time.sleep(0.5)
            self.serial.write(b'!')
            time.sleep(0.1)
            self.serial.read(256)  # drain
            self.connected = True
        except serial.SerialException as e:
            if self.serial:
                try:
                    self.serial.close()
                except serial.SerialException:
                    pass
                self.serial = None
            raise ConnectionError(f"Failed to connect: {e}") from e

    def disconnect(self) -> None:
        if self.serial and self.serial.is_open:
            self.serial.close()
        self.serial = None
        self.connected = False

    def read_data(self, addr: int, size: int) -> bytes:
        # A{addr}\0, R, read 64-byte chunks with '1' to continue, '0' to stop
        if self.serial is None:
            raise RuntimeError("Not connected")

        for c in f"A{addr:x}\0":
            self.serial.write(c.encode('ascii'))
            time.sleep(0.005)
        time.sleep(0.05)

        self.serial.write(b'R')
        time.sleep(0.1)

        chunks: list[bytes] = []
        received = 0
        while received < size and (chunk := self.serial.read(64)):
            chunks.append(chunk)
            received += len(chunk)
            if received < size:
                self.serial.write(b'1')
                time.sleep(0.02)

        self.serial.write(b'0')
        time.sleep(0.05)
        return b''.join(chunks)

    def write_data(self, addr: int, data: bytes) -> int:
        # A{addr}\0, then 'W' + 64-byte chunks, '1' ack between each
        if self.serial is None:
            raise RuntimeError("Not connected")

        for c in f"A{addr:x}\0":
            self.serial.write(c.encode('ascii'))
            time.sleep(0.005)
        time.sleep(0.05)

        for chunk in batched(data, 64):
            self.serial.write(b'W' + bytes(chunk).ljust(64, b'\xff'))
            # Wait for device to acknowledge with '1'
            ack = self.serial.read(1)
            if ack != b'1':
                break

        return len(data)

    def soft_reconnect(self) -> bool:
        if not self.serial or not self.serial.is_open:
            return False

        port = self.port
        try:
            self.serial.close()
            time.sleep(0.5)
            self.serial = serial.Serial(
                port=port,
                baudrate=BAUDRATE,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=TIMEOUT,
            )
            self.serial.dtr = True
            self.serial.rts = True
            time.sleep(0.5)
            self.serial.reset_input_buffer()
            self.serial.reset_output_buffer()
            self.serial.write(b'!')
            time.sleep(0.1)
            self.serial.read(256)
            return True
        except Exception:
            return False

    def read_bitmap(self) -> bytes | None:
        try:
            bitmap = self.read_data(0, BITMAP_SIZE)
            if len(bitmap) >= BITMAP_SIZE:
                return bitmap
        except serial.SerialException:
            pass
        return None

    def get_filled_slots(self, bitmap: bytes | None = None) -> list[int]:
        # bit=0 means filled
        if bitmap is None:
            bitmap = self.read_bitmap()
        if not bitmap:
            return []

        return [
            byte_idx * 8 + bit
            for byte_idx, byte_val in enumerate(bitmap)
            for bit in range(8)
            if not (byte_val & (0x80 >> bit))
        ]

    def read_image_data(self, slot_index: int, max_retries: int = 3) -> bytes:
        if slot_index < 0:
            raise ValueError("Invalid slot index")

        if self.serial is None:
            raise RuntimeError("Not connected")
        addr = 0x18 + slot_index

        data = b''
        for attempt in range(max_retries):
            self.serial.reset_input_buffer()
            self.serial.reset_output_buffer()
            time.sleep(0.05)

            data = self.read_data(addr, IMAGE_DATA_SIZE)
            if len(data) >= IMAGE_DATA_SIZE:
                return data

            if attempt < max_retries - 1:
                self.soft_reconnect()
                time.sleep(0.3)

        return data

    def erase_bitmap(self) -> bool:
        # 'k' = firmware erase-first-block (clears MFT, data remains)
        if self.serial is None:
            raise RuntimeError("Not connected")

        try:
            self.soft_reconnect()
            time.sleep(0.3)
            self.serial.write(b'k')
            # Firmware acknowledges with '1' when erase is complete
            ack = self.serial.read(1)
            return ack == b'1'
        except Exception:
            return False


def decode_2bpp_tile(tile_data: bytes) -> list[list[int]]:
    pixels = []
    for row in range(8):
        lo, hi = tile_data[row * 2], tile_data[row * 2 + 1]
        pixels.append([
            ((hi >> bit) & 1) << 1 | ((lo >> bit) & 1)
            for bit in range(7, -1, -1)
        ])
    return pixels


def decode_gb_camera_image(
    raw_data: bytes,
    palette: Palette | None = None,
) -> Image.Image:
    if palette is None:
        palette = PALETTE_GRAYSCALE

    img = Image.new('RGB', (IMAGE_WIDTH, IMAGE_HEIGHT))
    px = img.load()

    for tile_idx in range(TILES_X * TILES_Y):
        offset = tile_idx * TILE_SIZE
        if offset + TILE_SIZE > len(raw_data):
            break

        tile = decode_2bpp_tile(raw_data[offset:offset + TILE_SIZE])
        base_x = (tile_idx % TILES_X) * 8
        base_y = (tile_idx // TILES_X) * 8

        for y, row in enumerate(tile):
            for x, value in enumerate(row):
                px[base_x + x, base_y + y] = palette[value]

    return img


def _get_ffmpeg() -> str:
    from imageio_ffmpeg import get_ffmpeg_exe
    return get_ffmpeg_exe()


def create_gif(
    image_files: list[str | Path],
    output_path: str | Path,
    fps: int = 3,
    loop: int = 0,
) -> Path:
    if not image_files:
        raise ValueError("No images to convert")

    images = []
    for f in sorted(image_files):
        img = Image.open(f)
        images.append(img.copy())
        img.close()

    output = Path(output_path)
    images[0].save(
        output,
        save_all=True,
        append_images=images[1:],
        duration=1000 // fps,
        loop=loop,
    )
    return output


def create_mkv(
    image_dir: str | Path,
    output_path: str | Path,
    fps: int = 3,
    metadata: dict[str, str] | None = None,
) -> Path:
    ffmpeg = _get_ffmpeg()

    pattern = str(Path(image_dir) / '%06d.bmp')
    output = Path(output_path)
    cmd = [
        ffmpeg, '-y',
        '-framerate', str(fps),
        '-i', pattern,
        '-c:v', 'libx264',
        '-pix_fmt', 'yuv420p',
        '-crf', '18',
    ]

    if metadata:
        for key, value in metadata.items():
            if value:
                cmd.extend(['-metadata', f'{key}={value}'])

    cmd.append(str(output))

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"ffmpeg failed: {result.stderr}")
    return output
