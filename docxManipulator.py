import os
import re
import shutil
import subprocess
import logging

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


class DocxManipulator:
    def __init__(self, jar_name="sage-auto-styler.jar"):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.jar_file = os.path.join(script_dir, "DocxManipulator", jar_name)

        if not os.path.exists(self.jar_file):
            raise FileNotFoundError(f"JAR file not found: {self.jar_file}")

    @staticmethod
    def _needs_ascii_alias(path):
        """The JAR (Java on Windows) can mishandle command-line filenames that
        contain non-ASCII characters (e.g. the dotless 'ı', U+0131) or
        whitespace, reporting 'Input docx file location given not found'.
        Such names must be routed through a sanitized ASCII copy."""
        name = os.path.basename(path)
        return any(ord(c) > 127 for c in name) or any(c.isspace() for c in name)

    def docx_processor(self, docx_file, extra_args=None):
        # Validate input file
        docx_file = os.path.normpath(docx_file)

        if not os.path.isfile(docx_file):
            logger.error(f"Input file not found: {docx_file}")
            return False, None

        if not docx_file.lower().endswith(".docx"):
            logger.error(f"Input file is not a .docx file: {docx_file}")
            return False, None

        # Build output filename
        as_docx = re.sub(r"\.docx$", "_AS.docx", docx_file, flags=re.IGNORECASE)

        # Route risky (non-ASCII / spaced) filenames through a sanitized ASCII
        # copy so the JAR can locate the input, then map its "_AS" output back.
        jar_input = docx_file
        alias_path = None
        if self._needs_ascii_alias(docx_file):
            alias_path = os.path.join(
                os.path.dirname(docx_file), f"_jaras_{os.getpid()}.docx"
            )
            try:
                shutil.copy2(docx_file, alias_path)
                jar_input = alias_path
                logger.info(f"Risky filename — running JAR on ASCII alias: {alias_path}")
            except OSError as exc:
                logger.warning(f"Could not create ASCII alias ({exc}); "
                               f"running on original name.")
                alias_path = None
                jar_input = docx_file

        # Build and run command
        command = ["java", "-jar", self.jar_file, "-dx", jar_input, "-ipas"]
        if extra_args:
            command.extend(extra_args)
        logger.info(f"Running command: {' '.join(command)}")

        try:
            try:
                process = subprocess.run(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    universal_newlines=True,
                    timeout=300  # 5-minute timeout to prevent hangs
                )
            except FileNotFoundError:
                logger.error("Java is not installed or not found in PATH.")
                return False, None
            except subprocess.TimeoutExpired:
                logger.error("Process timed out after 300 seconds.")
                return False, None

            # Log output for debugging
            logger.debug(f"Return code: {process.returncode}")
            logger.debug(f"STDOUT:\n{process.stdout}")
            if process.stderr:
                logger.warning(f"STDERR:\n{process.stderr}")

            # Validate result
            if process.returncode != 0:
                logger.error(f"Process exited with return code {process.returncode}")
                return False, None

            if not re.search(r"Process Completed Successfully", process.stdout, re.IGNORECASE):
                logger.error("Success message not found in process output.")
                return False, None

            # Map the alias's "_AS" output back to the expected as_docx name.
            if alias_path is not None:
                alias_stem = re.sub(r"\.docx$", "", os.path.basename(alias_path),
                                    flags=re.IGNORECASE)
                alias_as = os.path.join(
                    os.path.dirname(docx_file), f"{alias_stem}_AS.docx"
                )
                if os.path.isfile(alias_as):
                    try:
                        if os.path.exists(as_docx):
                            os.remove(as_docx)
                        os.rename(alias_as, as_docx)
                    except OSError as exc:
                        logger.error(f"Failed to map alias output to {as_docx}: {exc}")
                        return False, None
        finally:
            if alias_path is not None and os.path.isfile(alias_path):
                try:
                    os.remove(alias_path)
                except OSError:
                    pass

        if not os.path.isfile(as_docx):
            logger.error(f"Expected output file not found: {as_docx}")
            return False, None

        logger.info(f"Processing completed successfully: {as_docx}")
        return True, as_docx

    def docx_preclean(self, docx_file, jar_args=None, timeout=300, success_pattern=None):
        """Run pre-clean on a docx file using the local JAR.

        The JAR is invoked with the ``-pre`` flag and produces a file named
        ``<stem>_PRE.docx`` alongside the original.  On success the method:

        1. Moves the **original** file into an ``org/`` sub-folder that sits
           next to the original file (creating it if necessary).
        2. Renames ``<stem>_PRE.docx`` → ``<stem>.docx`` so downstream code
           sees the same filename it started with.

        Args:
            docx_file: Absolute (or relative) path to the input ``.docx`` file.
            jar_args: Additional arguments passed to the JAR after
                ``-dx <file>``.  Defaults to ``["-pre"]`` which matches the
                command-line documented in the YAML config.
            timeout: Subprocess timeout in seconds (default 300).
            success_pattern: Regex matched against stdout to confirm the JAR
                succeeded.  Defaults to ``"Process Completed Successfully"``.

        Returns:
            ``(True, final_docx_path)`` on success, ``(False, None)`` on any
            failure.  *final_docx_path* is the renamed ``<stem>.docx`` in its
            original directory.
        """
        docx_file = os.path.normpath(docx_file)

        # ── input validation ────────────────────────────────────────────────
        if not os.path.isfile(docx_file):
            logger.error(f"Input file not found: {docx_file}")
            return False, None
        if not docx_file.lower().endswith(".docx"):
            logger.error(f"Input file is not a .docx file: {docx_file}")
            return False, None

        if jar_args is None:
            jar_args = ["-pre"]
        if success_pattern is None:
            success_pattern = r"Process Completed Successfully"

        # ── derive path names ───────────────────────────────────────────────
        parent_dir = os.path.dirname(docx_file)
        basename   = os.path.basename(docx_file)               # e.g. "report.docx"
        stem       = re.sub(r"\.docx$", "", basename, flags=re.IGNORECASE)

        pre_docx   = os.path.join(parent_dir, f"{stem}_PRE.docx")  # JAR output
        org_dir    = os.path.join(parent_dir, "org")                # archive folder
        org_file   = os.path.join(org_dir, basename)                # original's new home

        # ── run the JAR ─────────────────────────────────────────────────────
        # Route risky (non-ASCII / spaced) filenames through a sanitized ASCII
        # copy so the JAR can locate the input, then map its "_PRE" output back
        # to the expected name.
        jar_input = docx_file
        alias_path = None
        if self._needs_ascii_alias(docx_file):
            alias_path = os.path.join(parent_dir, f"_jarpre_{os.getpid()}.docx")
            try:
                shutil.copy2(docx_file, alias_path)
                jar_input = alias_path
                logger.info(f"Risky filename — running JAR on ASCII alias: {alias_path}")
            except OSError as exc:
                logger.warning(f"Could not create ASCII alias ({exc}); "
                               f"running on original name.")
                alias_path = None
                jar_input = docx_file

        command = ["java", "-jar", self.jar_file, "-dx", jar_input] + list(jar_args)
        logger.info(f"Running pre-clean command: {' '.join(command)}")

        try:
            try:
                process = subprocess.run(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    universal_newlines=True,
                    timeout=timeout,
                )
            except FileNotFoundError:
                logger.error("Java is not installed or not found in PATH.")
                return False, None
            except subprocess.TimeoutExpired:
                logger.error(f"Pre-clean timed out after {timeout} seconds.")
                return False, None

            logger.debug(f"Return code: {process.returncode}")
            logger.debug(f"STDOUT:\n{process.stdout}")
            if process.stderr:
                logger.warning(f"STDERR:\n{process.stderr}")

            if process.returncode != 0:
                logger.error(f"Pre-clean exited with return code {process.returncode}")
                return False, None

            if not re.search(success_pattern, process.stdout, re.IGNORECASE):
                logger.error("Success message not found in pre-clean output.")
                return False, None

            # Map the alias's "_PRE" output back to the expected pre_docx name.
            if alias_path is not None:
                alias_stem = re.sub(r"\.docx$", "", os.path.basename(alias_path),
                                    flags=re.IGNORECASE)
                alias_pre = os.path.join(parent_dir, f"{alias_stem}_PRE.docx")
                if os.path.isfile(alias_pre):
                    try:
                        if os.path.exists(pre_docx):
                            os.remove(pre_docx)
                        os.rename(alias_pre, pre_docx)
                    except OSError as exc:
                        logger.error(f"Failed to map alias output to {pre_docx}: {exc}")
                        return False, None
        finally:
            if alias_path is not None and os.path.isfile(alias_path):
                try:
                    os.remove(alias_path)
                except OSError:
                    pass

        if not os.path.isfile(pre_docx):
            logger.error(f"Expected pre-clean output not found: {pre_docx}")
            return False, None

        # ── move original → org/ ────────────────────────────────────────────
        try:
            os.makedirs(org_dir, exist_ok=True)
            shutil.move(docx_file, org_file)
            logger.info(f"Original file archived: {org_file}")
        except OSError as exc:
            logger.error(f"Failed to move original file to org/: {exc}")
            return False, None

        # ── rename _PRE.docx → original name ───────────────────────────────
        try:
            os.rename(pre_docx, docx_file)
            logger.info(f"Pre-cleaned file promoted to: {docx_file}")
        except OSError as exc:
            logger.error(f"Failed to rename pre-cleaned file: {exc}")
            # Attempt to restore the original so nothing is left in a broken state
            try:
                shutil.move(org_file, docx_file)
                logger.warning("Restored original file after rename failure.")
            except OSError as restore_exc:
                logger.error(f"Restore also failed — manual intervention required: {restore_exc}")
            return False, None

        logger.info(f"Pre-clean completed successfully: {docx_file}")
        return True, docx_file


# if __name__ == "__main__":
#     docx_input = r"V:\FOR_BREAKDOWN\ParaStyler_INPUT\SAGE\PIN_1428793\PIN_1428793_CLN.docx"
#
#     try:
#         manipulator = DocxManipulator()
#         success, output_file = manipulator.docx_preclean(docx_input)
#
#         if success:
#             print(f"Output file: {output_file}")
#         else:
#             print("Processing failed. Check logs above for details.")
#     except FileNotFoundError as e:
#         print(e)
