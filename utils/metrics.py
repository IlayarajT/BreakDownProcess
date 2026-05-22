import time
import logging

log = logging.getLogger(__name__)


class Metrics:
    def __init__(self):
        self.timings = {}

    def track(self, name):
        def wrapper(func):
            def inner(*args, **kwargs):
                start = time.time()
                result = func(*args, **kwargs)
                elapsed = time.time() - start
                self.timings.setdefault(name, []).append(elapsed)
                log.debug("Metric %s: %.2fs", name, elapsed)
                return result
            return inner
        return wrapper

    def summary(self):
        return {
            k: round(sum(v) / len(v), 2)
            for k, v in self.timings.items()
        }


metrics = Metrics()
