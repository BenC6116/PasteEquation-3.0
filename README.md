# PasteEquation v3.0.1 - Improved Version

## 🔧 What's New in This Version

This is an **improved version** of the original [PasteEquation by Foxxey](https://github.com/Foxxey/PasteEquation) with critical stability fixes.

### 🐛 Fixed: Memory Crashes
- **Problem**: Original version would crash Word with white screen when pasting large ChatGPT responses (20+ equations)
- **Solution**: Implemented throttled paste processing that handles equations in chunks
- **Result**: Stable performance even with 50+ equations

## 📋 About PasteEquation

This Microsoft Word add-in allows you to paste multiline text with equations from ChatGPT directly into MS Word on Windows. The equations are automatically converted to proper Word equation format.

**Original Project**: [PasteEquation by Foxxey](https://github.com/Foxxey/PasteEquation)  
**Related Project**: [CopyEquation (Improved Version)](https://github.com/BenC6116/CopyEquation) - Browser extension for copying from ChatGPT

## 📦 Installation

1. **Download**: Go to the [Releases tab](../../releases) and download `setup.zip`
2. **Uninstall**: Remove any previous version of PasteEquation first
3. **Install**: Extract and run `setup.exe`
4. **Restart**: Close and reopen Microsoft Word
5. **Verify**: Check File → Options → Add-ins → COM Add-ins to ensure PasteEquation is enabled

## 🚀 How It Works

1. Copy text with equations from ChatGPT (use the [CopyEquation browser extension](https://github.com/BenC6116/CopyEquation))
2. Place cursor in Word document
3. Press **Ctrl+V** to paste
4. Equations are automatically converted to Word's native equation format

## 🔧 Technical Improvements

- **Chunked Processing**: Equations processed in batches of 10
- **Memory Management**: 200ms pauses prevent Word's OMML engine overflow
- **Stability**: No more crashes on large equation sets
- **Compatibility**: Works with original CopyEquation workflow

## 🙏 Credits

- **Original Author**: [Foxxey](https://github.com/Foxxey) - Created the original PasteEquation
- **This Version**: Stability improvements and crash fixes

## 📄 License

This project maintains the same license as the original PasteEquation project.