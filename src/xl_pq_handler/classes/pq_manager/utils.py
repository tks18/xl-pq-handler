# utils.py
import logging


def get_logger(name: str):
    """Configures and returns a standard logger."""
    logging.basicConfig(
        level=logging.INFO,
        format="[%(levelname)s] (%(name)s) %(message)s"
    )
    return logging.getLogger(name)
