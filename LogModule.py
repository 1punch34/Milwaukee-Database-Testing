import logging

class Logger:
    """
    Set up a logger with console and file handlers.
    
    Args:
        name (str): The name of the logger.
    """

    def __init__(self, name: str) -> None:
        self.name = name
        self.logger = self.setup_logger()

    def setup_logger(self) -> logging.Logger:
        logger = logging.getLogger(self.name)
        logger.setLevel(logging.INFO)

        # Avoid adding handlers multiple times
        if not logger.hasHandlers():
            # Formatter
            formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")

            # Console handler
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)

            # File handler for general logs
            file_handler = logging.FileHandler("general.log")
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)

            # File handler for error logs only
            error_handler = logging.FileHandler("errors.log")
            error_handler.setFormatter(formatter)
            error_handler.setLevel(logging.ERROR)
            logger.addHandler(error_handler)

        return logger