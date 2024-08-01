import logging
from tabulate import tabulate
import time
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
            # Formatter with line number
            formatter = logging.Formatter(
                "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
            )

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

    def errortrace(logger: logging.Logger, exception: Exception):
        """Log an exception with a stack trace."""
        logger.error("An error occurred", exc_info=True)

class LogData:
    inventoryCount = 0
    productDetailsCount = 0
    uomCount = 0
    priceCount = 0
    errorList = []
    startTime = None

    def startTimer(cls):
        cls.startTime = time.time()
    @classmethod
    def getElapsedTime(cls):
        if cls.startTime is not None:
            elapsed_time = time.time() - cls.startTime
            return elapsed_time / 60  # Convert seconds to minutes
        else:
            return None  # Timer hasn't started

    @classmethod
    def addInventoryCount(cls):
        cls.inventoryCount += 1

    @classmethod
    def addProductDetailsCount(cls):
        cls.productDetailsCount += 1

    @classmethod
    def addUOMCount(cls):
        cls.uomCount += 1

    @classmethod
    def addPriceCount(cls):
        cls.priceCount += 1

    @classmethod
    def logError(cls, error_message):
        cls.errorList.append(error_message)

    @classmethod
    def getLogData(cls):
        elapsed_time = cls.getElapsedTime()
        elapsed_time_str = f"{elapsed_time:.2f} minutes" if elapsed_time is not None else "Timer not started"

        data = [["Elapsed Time", elapsed_time_str],
            ["Inventory Records Created", cls.inventoryCount],
            ["ProductDetails Records Created", cls.productDetailsCount],
            ["UOM Records Created", cls.uomCount],
            ["Price Records Created", cls.priceCount],
            ["Errors logged", len(cls.errorList)]
            
        ]

        table = tabulate(data, tablefmt="grid")
        return table