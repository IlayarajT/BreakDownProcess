import logging
import psutil

class ResourceMonitor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def log_resource_usage(self):
        self.logger.info(f"Memory usage: {psutil.virtual_memory().percent}%")
        self.logger.info(f"CPU usage: {psutil.cpu_percent(interval=1)}%")
        self.logger.info(f"Open files: {len(psutil.Process().open_files())}")
