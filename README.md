# iPhone Photo Copier

A Python program to copy all photos from iPhone to Windows.

## Requirements

- Windows 10 or Windows 11
- Python 3.x
- pywin32 library: `pip install pywin32`
- iPhone connected via USB with authorization granted

## Usage

### Method 1: Batch file (Windows)
Double-click the `iphone_photo_copier.bat` file

### Method 2: Command line
```bash
python main.py
```

The program will automatically copy all photos from the iPhone to the configured destination folder (default: `Desktop/iphone_photo_copier`).

## Configuration

The `config.json` file allows you to customize the destination folder:

```json
{
  "use_desktop": true,
  "custom_path": "C:/Users/username/Pictures/iPhone"
}
```

- If `use_desktop` is `true`, photos are copied to `Desktop/iphone_photo_copier`
- If `use_desktop` is `false`, the path in `custom_path` is used

## Note

**This program is designed to copy all photos in a single session**.

### Recommendations

1. Do not disconnect the iPhone during copying
2. To maximize speed, complete the entire copy in a single session

### MTP Cache Issue

When reconnecting the iPhone after disconnecting:
- Windows loses the MTP device cache
- The program must re-scan all folders from the beginning
- Even already copied folders are analyzed slowly
- The operation that was fast (thanks to cache) becomes slow again


## Features

- ✅ Automatic copy of all photos organized by folders (rolls)
- ✅ Automatic skip of already copied files
- ✅ Detailed log with timestamps
- ✅ Estimated time remaining
- ✅ Final statistics (files copied, average speed, total time)

## Output

The program creates:
- One folder for each iPhone "roll"
- A `log.txt` file with details of all operations

## Technical Notes

The program uses the MTP (Media Transfer Protocol) via Windows COM to access iPhone files. This is the only way to access iPhone photos on Windows without additional software.