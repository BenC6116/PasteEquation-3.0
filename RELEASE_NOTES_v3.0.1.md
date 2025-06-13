# PasteEquation v3.0.1 - Memory Crash Fix

## 🐛 Bug Fixes

### Fixed Memory Crashes with Large Equation Sets
- **Issue**: Word would crash with white screen when pasting responses containing many equations (typically 20+ equations)
- **Root Cause**: Word's OMML (Office Math Markup Language) engine would run out of memory when processing too many equations at once
- **Solution**: Implemented throttled paste loop that processes equations in chunks

### Technical Details
- Equations are now processed in batches of 10
- 200ms pause between batches allows Word's memory management to catch up
- Prevents Word from exceeding its 2GB address space limit
- Works reliably even on systems with 4GB RAM

## 🚀 Improvements
- More stable performance with large ChatGPT responses
- No functional changes to the core equation conversion
- Maintains full compatibility with existing workflows

## 📦 Installation
1. Uninstall the previous version of PasteEquation
2. Run `setup.exe` from the downloaded ZIP file
3. Restart Microsoft Word
4. Verify the add-in is enabled in File → Options → Add-ins → COM Add-ins

## 🧪 Testing
- Tested with responses containing 50+ equations
- Memory usage remains stable throughout paste operation
- No more white-screen crashes during large paste operations

---

**Note**: This is a maintenance release focused on stability. The core functionality and user experience remain unchanged.
