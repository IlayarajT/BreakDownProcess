import os
import re
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

    def docx_processor(self, docx_file):
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

        # Build and run command
        command = ["java", "-jar", self.jar_file, "-dx", docx_file, "-ipas"]
        logger.info(f"Running command: {' '.join(command)}")

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

        if not os.path.isfile(as_docx):
            logger.error(f"Expected output file not found: {as_docx}")
            return False, None

        logger.info(f"Processing completed successfully: {as_docx}")
        return True, as_docx


# if __name__ == "__main__":
#     docx_input = r"V:\FOR_BREAKDOWN\ParaStyler_INPUT\SAGE\PIN_1428793\PIN_1428793_CLN.docx"
#
#     try:
#         manipulator = DocxManipulator()
#         success, output_file = manipulator.docx_processor(docx_input)
#
#         if success:
#             print(f"Output file: {output_file}")
#         else:
#             print("Processing failed. Check logs above for details.")
#     except FileNotFoundError as e:
#         print(e)