from word_failures import WordFailureType
import pythoncom
import time
import win32com.client
import traceback
from pathlib import Path
import os
import subprocess


def restart_word():
    """Create a fresh Word instance with proper cleanup."""
    import win32com.client
    import pythoncom

    print("[INFO] Restarting Word application...")

    # Try to clean up existing instances
    for _ in range(3):
        try:
            # Kill any running WINWORD processes
            subprocess.run(
                "taskkill /f /im WINWORD.EXE",
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=5
            )
        except Exception:
            pass

        time.sleep(2)  # Wait for processes to die

        try:
            # Clean up COM
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

            # Re-initialize COM
            pythoncom.CoInitialize()

            # Create new instance with optimizations for large files
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            word.AutomationSecurity = 3
            word.ScreenUpdating = False  # Disable screen updates for speed

            # Optimize for performance
            word.Options.SavePropertiesPrompt = False
            word.Options.SaveNormalPrompt = False
            word.Options.UpdateLinksAtOpen = False
            word.Options.CheckGrammarAsYouType = False
            word.Options.CheckSpellingAsYouType = False

            # Give Word time to initialize
            for i in range(10):
                try:
                    # Test if Word is responsive
                    _ = word.Version
                    print(f"[SUCCESS] Word restarted successfully (Version: {word.Version})")
                    return word
                except Exception:
                    if i < 9:
                        time.sleep(1)
                    else:
                        raise

        except Exception as e:
            print(f"[WARN] Word restart attempt failed: {e}")
            time.sleep(3)
            continue
    raise Exception("Failed to restart Word after multiple attempts")



def open_word_document(word, file_name, retries=3, delay_between_retries=5, timeout=60):
    """
    Safely open a Word document with retry logic and timeout handling.

    Args:
        word: Word.Application instance
        file_name: Path to document
        retries: Number of retry attempts
        delay_between_retries: Delay between retries in seconds
        timeout: Maximum seconds to wait for Word to respond

    Returns:
        tuple: (document, failure_type, word_instance)
    """
    last_exception = None

    for attempt in range(1, retries + 1):
        try:
            print(f"[INFO] Opening document attempt {attempt}/{retries}: {Path(file_name).name}")

            # Prepare file path (remove long path prefix for COM)
            com_path = file_name.replace('\\\\?\\', '') if file_name.startswith('\\\\?\\') else file_name

            # For large files, increase timeout
            file_size = os.path.getsize(com_path) if os.path.exists(com_path) else 0
            file_timeout = timeout
            if file_size > 50 * 1024 * 1024:  # > 50MB
                file_timeout = timeout * 2  # Double timeout for large files
                print(
                    f"[INFO] Large file detected ({file_size / 1024 / 1024:.1f}MB), using extended timeout: {file_timeout}s")

            # Set a timeout for the COM operation
            start_time = time.time()

            # Try to open with full access
            try:
                # Set Word to be more responsive
                word.DisplayAlerts = 0
                word.ScreenUpdating = False
                word.EnableCancelKey = 0  # Disable ESC key interrupt

                # Open document with extended timeout.
                # F1: OpenAndRepair=False — using True on a network/mapped drive
                # causes Word to open the file as "Document1" (an internal repair
                # copy), which then makes any doc.Save() call trigger a Save As
                # dialog instead of saving silently.
                doc = word.Documents.Open(
                    FileName=com_path,
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Revert=True,
                    NoEncodingDialog=True,
                    OpenAndRepair=False,
                    PasswordDocument="",
                    WritePasswordDocument="",
                    Visible=True
                )

                # Wait for document to fully load
                load_wait_start = time.time()
                while time.time() - load_wait_start < file_timeout:
                    try:
                        # Try to access a property to see if document is ready
                        _ = doc.Name
                        print(f"[INFO] Document loaded in {time.time() - start_time:.1f}s")
                        break
                    except Exception:
                        if time.time() - load_wait_start > file_timeout:
                            raise TimeoutError(f"Document took too long to load (> {file_timeout}s)")
                        time.sleep(1)
                        pythoncom.PumpWaitingMessages()  # Keep COM alive

            except TimeoutError as toe:
                print(f"[WARN] Timeout opening document: {toe}")
                # Try to close the document if it partially opened
                try:
                    if 'doc' in locals():
                        doc.Close(SaveChanges=False)
                except Exception:
                    pass
                raise

            # ========== FIXED PROTECTION CHECK ==========
            # Check if document is protected BEFORE trying to unprotect
            try:
                if hasattr(doc, "ProtectionType"):
                    protection_type = doc.ProtectionType
                    # wdNoProtection = -1
                    if protection_type != -1:
                        print(f"[INFO] Document is protected (Type: {protection_type}), attempting to unprotect...")
                        try:
                            doc.Unprotect()
                            print(f"[INFO] Document unprotected successfully")
                        except Exception as unprotect_error:
                            error_msg = str(unprotect_error).lower()
                            if "already unprotected" not in error_msg:
                                print(f"[WARN] Could not unprotect: {unprotect_error}")
            except Exception as e:
                error_msg = str(e).lower()
                if "already unprotected" not in error_msg:
                    print(f"[INFO] Protection check: {e}")
            # ========== END FIX ==========

            # Re-enable screen updating for macros
            word.ScreenUpdating = True

            print(f"[SUCCESS] Document opened successfully")
            return doc, None, word

        except pythoncom.com_error as e:
            # Handle COM errors specifically
            last_exception = e
            hresult = e.hresult if hasattr(e, 'hresult') else -1

            print(f"[COM ERROR] HRESULT: {hresult:#010x}, Message: {e}")

            # Known COM errors and their handling
            com_errors = {
                -2147023179: "INTERFACE_UNKNOWN",  # 0x800706B5
                -2147023174: "RPC_SERVER_UNAVAILABLE",  # 0x800706BA
                -2147220995: "SERVER_NOT_CONNECTED",  # 0x800401FD
                -2147418111: "CALL_REJECTED",  # 0x80010001
                -2147418110: "CALL_CANCELED",  # 0x80010002
                -2147417856: "SERVER_BUSY",  # 0x80010010
                -2147023170: "RPC_SERVER_TOO_BUSY",  # 0x800706BE
                -2146959355: "SERVER_EXEC_FAILURE",  # 0x80080005
            }

            error_name = com_errors.get(hresult, "UNKNOWN_COM_ERROR")
            print(f"[INFO] COM error type: {error_name}")

            # Different handling based on error type
            if hresult in [-2147023174, -2147220995, -2146959355]:  # RPC/SERVER errors
                print(f"[INFO] Restarting Word due to {error_name}...")
                try:
                    word = restart_word()
                except Exception as restart_error:
                    print(f"[ERROR] Failed to restart Word: {restart_error}")

                if attempt < retries:
                    print(f"[INFO] Waiting {delay_between_retries}s before retry...")
                    time.sleep(delay_between_retries)
                    continue

            elif hresult in [-2147418111, -2147417856, -2147023170]:  # BUSY/CALL_REJECTED errors
                print(f"[INFO] Word is busy, waiting and retrying...")
                pythoncom.PumpWaitingMessages()
                if attempt < retries:
                    # Increase wait time for busy errors
                    wait_time = delay_between_retries * attempt
                    print(f"[INFO] Waiting {wait_time}s...")
                    time.sleep(wait_time)
                    continue

            elif hresult == -2147023179:  # INTERFACE_UNKNOWN
                print(f"[INFO] COM interface issue, restarting Word...")
                try:
                    word = restart_word()
                except Exception:
                    pass

                if attempt < retries:
                    time.sleep(delay_between_retries)
                    continue

            # If we get here, treat as general error
            if attempt < retries:
                print(f"[INFO] Retrying in {delay_between_retries}s...")
                time.sleep(delay_between_retries)
                continue

        except Exception as e:
            last_exception = e
            error_msg = str(e).lower()
            print(f"[WARN] Open attempt {attempt} failed: {error_msg}")

            # Analyze error type
            if "corrupt" in error_msg or "repair" in error_msg:
                print(f"[ERROR] Document appears corrupted")
                return None, WordFailureType.CORRUPTED, word

            elif "cannot find" in error_msg or "file not found" in error_msg or "does not exist" in error_msg:
                print(f"[ERROR] File not found")
                return None, WordFailureType.FILE_NOT_FOUND, word

            elif "protected view" in error_msg or "blocked" in error_msg or "security" in error_msg:
                print(f"[INFO] Document blocked by Protected View")
                if attempt < retries:
                    time.sleep(delay_between_retries)
                    continue

            elif "timeout" in error_msg:
                print(f"[INFO] Operation timed out")
                if attempt < retries:
                    # Increase timeout for next attempt
                    time.sleep(delay_between_retries * 2)
                    continue

            elif "permission" in error_msg or "access denied" in error_msg:
                print(f"[INFO] Permission error, trying read-only...")
                # Try read-only mode
                try:
                    com_path = file_name.replace('\\\\?\\', '') if file_name.startswith('\\\\?\\') else file_name
                    doc = word.Documents.Open(
                        FileName=com_path,
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Revert=True,
                        NoEncodingDialog=True,
                        OpenAndRepair=False,
                        Visible=True
                    )
                    return doc, WordFailureType.FILE_LOCKED, word
                except Exception:
                    if attempt < retries:
                        time.sleep(delay_between_retries)
                        continue

            else:
                # Generic error, retry
                if attempt < retries:
                    print(f"[INFO] Retrying after generic error...")
                    time.sleep(delay_between_retries)
                    continue

    # All retries failed
    print(f"[ERROR] Failed to open document after {retries} attempts")
    if last_exception:
        print(f"[ERROR] Last exception: {last_exception}")

    return None, WordFailureType.COM_FAILURE, word