import logging


# Logger Config
# Each Daily task has its own logger
# Make sure to create a logs folder

def setup_workilist_logger(name: str) -> logging.Logger:
    """
        Logger for the worklist.py
        Keeps records of logs
    """
    logging.basicConfig(
        level=logging.INFO, 
        format= "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.FileHandler("logs\\worklist_automation.log"),
            logging.StreamHandler()

        ]
    )
    

    return logging.getLogger(name)


def setup_rush_logger(name: str) -> logging.Logger:
    """
        Logger for the rush.py
        Keeps records of logs
    """
    logging.basicConfig(
        level=logging.INFO, 
        format= "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.FileHandler("logs\\rush_automation.log"),
            logging.StreamHandler()

        ]
    )
    

    return logging.getLogger(name)