# PicNRec Protocol Specification

A minimal specification for implementing PicNRec device communication.

## Hardware Requirements

### USB Connection
- **USB-A (PC side) to USB-C (device side) cable recommended**
- USB-C to USB-C can cause intermittent read failures on some systems (observed on Linux; may be fine elsewhere)
- Device: CH340 USB-to-Serial (VID: 0x1A86, PID: 0x7523)
- Linux device path: `/dev/ttyUSB0` (typically)

### Serial Configuration
```
Baud rate:  1,000,000 (1 Mbps)
Data bits:  8
Parity:     None
Stop bits:  1
DTR:        High (True)
RTS:        High (True)
Timeout:    2000ms recommended
```

## Protocol Commands

All commands are ASCII. The device operates in a request-response pattern.

| Command | Format | Description |
|---------|--------|-------------|
| Init | `!` | Reset/initialize device |
| Set Address | `A{hex}\0` | Set read address (hex string + null terminator) |
| Read | `R` | Start reading from current address |
| Continue | `1` | Request next 64-byte chunk |
| Stop | `0` | End read operation |

## Data Structures

### Address Space
```
Address 0x00:       Allocation bitmap (2340 bytes)
Address 0x18 + N:   Image slot N (where N = 0, 1, 2, ...)
```

### Allocation Bitmap
- 2340 bytes at address 0x00
- Each byte represents 8 image slots
- Bit ordering: MSB = first slot in group
- **Bit = 0: Slot contains image**
- **Bit = 1: Slot is empty**

Example:
```
Byte 0x00 = 0b00000000 → Slots 0-7 all filled
Byte 0x3F = 0b00111111 → Slots 0-1 filled, 2-7 empty
Byte 0xFF = 0b11111111 → All 8 slots empty
```

### Image Data
- 3584 bytes per image
- 128×112 pixels, 2 bits per pixel (4 colors)
- Organized as 8×8 pixel tiles (224 tiles total: 16 wide × 14 tall)
- 16 bytes per tile (2 bytes per row × 8 rows)

## Read Sequence

### 1. Connection Setup
```
1. Open serial port with settings above
2. Set DTR = true, RTS = true
3. Wait 500ms for connection to stabilize
4. Send '!' (init command)
5. Wait 100ms
6. Read and discard any response (drain buffer)
```

### 2. Reading Data
```
1. Send address: 'A' + hex_string + '\0'
   - Send each byte with 5ms delay between
   - Wait 50ms after complete address

2. Send 'R' to start read
   - Wait 100ms after R

3. Read loop:
   - Read 64 bytes (with timeout)
   - If more data needed AND got 64 bytes:
     - Send '1' to continue
     - Wait 20ms
   - Repeat until all data received or timeout

4. Send '0' to end read
   - Wait 50ms after stop
```

### 3. Reading Bitmap (to count images)
```
address = 0x00
size = 2340 bytes
data = read_data(address, size)

count = 0
for each byte in data:
    for bit in 7..0:
        if (byte & (0x80 >> bit)) == 0:
            count += 1
```

### 4. Reading an Image
```
slot_number = N  (0-based)
address = 0x18 + slot_number
size = 3584 bytes
image_data = read_data(address, size)
```

## Image Decoding

### Tile Layout
Tiles stored linearly, row-major order:
```
Tile index = tile_row * 16 + tile_col
Byte offset = tile_index * 16
```

### 2BPP Pixel Decoding
Each tile row is 2 bytes: [low_plane, high_plane]
```
for row in 0..8:
    low  = data[tile_offset + row * 2]
    high = data[tile_offset + row * 2 + 1]

    for bit in 0..8:
        shift = 7 - bit
        pixel = ((high >> shift) & 1) << 1 | ((low >> shift) & 1)
        // pixel is 0, 1, 2, or 3
```

### Color Palette
```
Pixel 0 → White  (255, 255, 255) or (0xFF)
Pixel 1 → Light  (170, 170, 170) or (0xAA)
Pixel 2 → Dark   (85, 85, 85)    or (0x55)
Pixel 3 → Black  (0, 0, 0)       or (0x00)
```

## Timing Summary

| Operation | Delay |
|-----------|-------|
| After port open | 500ms |
| After '!' init | 100ms |
| Between address bytes | 5ms |
| After address complete | 50ms |
| After 'R' command | 100ms |
| After '1' continue | 20ms |
| After '0' stop | 50ms |
| Between images (reconnect) | 200ms |

## Reliability Pattern

For maximum reliability, use fresh connection per image:
```
for each slot:
    wait(200ms)
    connection = open_and_init()
    data = read_image(connection, slot)
    connection.close()
    process(data)
```

This prevents state accumulation and ensures clean reads.

## Error Handling

If a read returns fewer than expected bytes:
1. Send '0' to stop current operation
2. Close serial connection
3. Wait 500ms
4. Reopen connection and retry

Retrying without closing/reopening won't work; the CH340 chip can get stuck and requires a port reset to recover.

## Pseudocode Implementation

```rust
fn read_image(port: &str, slot: u32) -> Result<Vec<u8>, Error> {
    // Open connection
    let mut serial = Serial::open(port, 1_000_000)?;
    serial.set_dtr(true);
    serial.set_rts(true);
    sleep(500ms);

    // Init
    serial.write(b"!")?;
    sleep(100ms);
    serial.drain();

    // Send address
    let addr = 0x18 + slot;
    let addr_cmd = format!("A{:x}\0", addr);
    for byte in addr_cmd.bytes() {
        serial.write(&[byte])?;
        sleep(5ms);
    }
    sleep(50ms);

    // Start read
    serial.write(b"R")?;
    sleep(100ms);

    // Read chunks
    let mut data = Vec::new();
    while data.len() < 3584 {
        let chunk = serial.read(64, timeout=2000ms)?;
        if chunk.is_empty() {
            break;
        }
        data.extend(chunk);
        if data.len() < 3584 {
            serial.write(b"1")?;
            sleep(20ms);
        }
    }

    // Stop
    serial.write(b"0")?;
    sleep(50ms);
    serial.close();

    Ok(data)
}
```

## Constants Reference

```
IMAGE_WIDTH      = 128 pixels
IMAGE_HEIGHT     = 112 pixels
TILES_X          = 16
TILES_Y          = 14
TILE_SIZE        = 16 bytes
IMAGE_SIZE       = 3584 bytes
BITMAP_SIZE      = 2340 bytes
MAX_IMAGES       = 18720
BITMAP_ADDRESS   = 0x00
IMAGE_BASE_ADDR  = 0x18
```
