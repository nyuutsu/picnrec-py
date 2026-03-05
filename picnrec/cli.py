#!/usr/bin/env python3
"""PicNRec CLI Tool"""

import argparse
import datetime
import os
import subprocess
import sys
import time
from collections.abc import Callable
from pathlib import Path

import serial.tools.list_ports

from picnrec.core import (
    BAUDRATE,
    IMAGE_DATA_SIZE,
    MAX_IMAGES,
    MIN_VALID_BYTES,
    PALETTES,
    PALETTE_GRAYSCALE,
    PicNRecDevice,
    create_gif,
    create_mkv,
    decode_gb_camera_image,
)


def export_images(
    device: PicNRecDevice,
    slots: list[int],
    output_dir: Path,
    palette_name: str = 'grayscale',
    number_padding: int = 6,
    show_progress: bool = True,
) -> list[Path]:
    palette = PALETTES.get(palette_name, PALETTE_GRAYSCALE)
    output_dir.mkdir(parents=True, exist_ok=True)

    exported: list[Path] = []
    failed: list[int] = []

    for i, slot in enumerate(slots):
        if show_progress:
            print(f"\rExporting image {i + 1}/{len(slots)}...", end='', flush=True)

        if i > 0:
            time.sleep(0.2)
            device.soft_reconnect()

        try:
            raw_data = device.read_image_data(slot)
            if len(raw_data) >= IMAGE_DATA_SIZE:
                non_ff = len(raw_data) - raw_data.count(b'\xff')
                if non_ff > MIN_VALID_BYTES:
                    img = decode_gb_camera_image(raw_data, palette)
                    filepath = output_dir / f"{str(slot).zfill(number_padding)}.bmp"
                    img.save(filepath, 'BMP')
                    exported.append(filepath)
                else:
                    failed.append(slot)
            else:
                failed.append(slot)
        except Exception as e:
            print(f"\nWarning: Failed to export slot {slot}: {e}")
            failed.append(slot)

    if show_progress:
        print(f"\nExported {len(exported)} images to {output_dir}")
        if failed:
            print(f"Skipped {len(failed)} empty/failed slots")

    return exported


def cmd_info(args: argparse.Namespace, device: PicNRecDevice) -> None:
    print(f"Device: {device.port}")
    print(f"Baud rate: {BAUDRATE}")

    try:
        filled_slots = device.get_filled_slots()
        count = len(filled_slots)
        print(f"Stored images: {count}")
        if count > 0:
            print(f"Slot range: {min(filled_slots)} - {max(filled_slots)}")
        usage_pct = (count / MAX_IMAGES) * 100
        print(f"Storage used: {count}/{MAX_IMAGES} slots ({usage_pct:.1f}%)")
    except Exception as e:
        print(f"Error reading device info: {e}")


def cmd_view(args: argparse.Namespace, device: PicNRecDevice) -> None:
    try:
        raw_data = device.read_image_data(args.index)
        palette = PALETTES.get(args.palette, PALETTE_GRAYSCALE)
        img = decode_gb_camera_image(raw_data, palette)
        temp_path = Path(f"/tmp/picnrec_preview_{args.index}.png")
        img.save(temp_path, 'PNG')
        print(f"Image saved to: {temp_path}")

        try:
            if sys.platform == 'darwin':
                subprocess.run(['open', str(temp_path)], check=False)
            elif sys.platform == 'win32':
                os.startfile(temp_path)
            else:
                subprocess.run(['xdg-open', str(temp_path)], check=False)
        except (FileNotFoundError, OSError):
            pass
    except Exception as e:
        print(f"Error viewing image: {e}")


def cmd_export(args: argparse.Namespace, device: PicNRecDevice) -> None:
    output_dir = Path(args.output) if args.output else Path(
        f"picnrec_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )

    filled = device.get_filled_slots()
    if not filled:
        print("No images found on device.")
        return

    if args.all:
        slots = filled
    else:
        start = args.start
        end = args.end if args.end is not None else start
        if end < start:
            print("Error: End index must be >= start index")
            return
        filled_set = set(filled)
        slots = [s for s in range(start, end + 1) if s in filled_set]
        if not slots:
            print(f"No filled slots in range {start}-{end}")
            return

    print(f"Exporting {len(slots)} images...")
    exported = export_images(device, slots, output_dir, args.palette, args.padding)

    if args.gif and exported:
        gif_path = output_dir.with_suffix('.gif')
        print(f"Creating GIF: {gif_path}")
        create_gif(exported, gif_path, args.fps)

    if args.mkv and exported:
        mkv_path = output_dir.with_suffix('.mkv')
        print(f"Creating MKV: {mkv_path}")
        create_mkv(output_dir, mkv_path, args.fps)


def cmd_erase(args: argparse.Namespace, device: PicNRecDevice) -> None:
    bitmap = device.read_bitmap()
    if bitmap:
        filled = device.get_filled_slots(bitmap)
        print(f"Device currently has {len(filled)} images stored.")
    else:
        print("Warning: Could not read current bitmap.")

    if not args.force:
        print()
        print("WARNING: This will erase all image references!")
        print("The device will begin overwriting images from slot 0.")
        print("Image data is NOT zeroed, just marked as available.")
        print()
        confirm = input("Type 'ERASE' to confirm: ")
        if confirm != 'ERASE':
            print("Aborted.")
            return

    print("Erasing bitmap (MFT)...")
    if device.erase_bitmap():
        print("Erase complete. All slots marked as empty.")
        bitmap = device.read_bitmap()
        if bitmap:
            filled = device.get_filled_slots(bitmap)
            print(f"Verification: {len(filled)} images now reported.")
    else:
        print("Erase may have failed. Check device connection.")


def cmd_list_ports(args: argparse.Namespace) -> None:
    ports = serial.tools.list_ports.comports()
    if not ports:
        print("No serial ports found.")
        return

    print("Available serial ports:")
    for port in ports:
        vid_pid = f" [{port.vid:04x}:{port.pid:04x}]" if port.vid else ""
        print(f"  {port.device}{vid_pid} - {port.description}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description='Game Boy Camera PicNRec interface',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Examples:
  %(prog)s info                          Show device info
  %(prog)s view 0                        View first image
  %(prog)s export --start 0 --end 100    Export images 0-100
  %(prog)s export --all --gif --fps 5    Export all + create GIF
  %(prog)s erase                         Erase all stored images""",
    )

    parser.add_argument('-p', '--port', help='Serial port (auto-detect if not specified)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')

    sub = parser.add_subparsers(dest='command', help='Command to run')

    sub.add_parser('info', help='Show device information')

    p_view = sub.add_parser('view', help='View a single image')
    p_view.add_argument('index', type=int, help='Image slot index')
    p_view.add_argument('--palette', choices=list(PALETTES.keys()),
                        default='grayscale', help='Color palette')

    p_export = sub.add_parser('export', help='Export images')
    p_export.add_argument('--start', '-s', type=int, default=0, help='Start slot index')
    p_export.add_argument('--end', '-e', type=int, help='End slot index')
    p_export.add_argument('--all', '-a', action='store_true', help='Export all filled slots')
    p_export.add_argument('--output', '-o', help='Output directory')
    p_export.add_argument('--palette', choices=list(PALETTES.keys()),
                          default='grayscale', help='Color palette')
    p_export.add_argument('--padding', type=int, default=6, help='Filename number padding')
    p_export.add_argument('--gif', action='store_true', help='Also create GIF')
    p_export.add_argument('--mkv', action='store_true', help='Also create MKV')
    p_export.add_argument('--fps', type=int, default=3, help='FPS for GIF/MKV')

    p_erase = sub.add_parser('erase', help='Erase stored images')
    p_erase.add_argument('--force', '-f', action='store_true', help='Skip confirmation')

    sub.add_parser('ports', help='List available serial ports')

    args = parser.parse_args()

    if args.command == 'ports':
        cmd_list_ports(args)
        return 0

    if not args.command:
        parser.print_help()
        return 1

    device = PicNRecDevice(port=args.port)
    try:
        print("Connecting to PicNRec...")
        device.connect()
        print(f"Connected to {device.port} at {BAUDRATE} baud")

        commands: dict[str, Callable] = {
            'info': cmd_info,
            'view': cmd_view,
            'export': cmd_export,
            'erase': cmd_erase,
        }
        if args.command in commands:
            commands[args.command](args, device)

    except ConnectionError as e:
        print(f"Connection error: {e}")
        return 1
    except KeyboardInterrupt:
        print("\nAborted.")
        return 130
    except Exception as e:
        print(f"Error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1
    finally:
        device.disconnect()

    return 0


if __name__ == '__main__':
    sys.exit(main())
