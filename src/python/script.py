import logging
import os
import pandas as pd
import time

def main(thread,input, output, dir):
    # Configure logging
    now = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    logging.basicConfig(
        filename=os.path.join(output, f'{now}.log'),  # Log file name
        level=logging.DEBUG,    # Log level
        format='%(asctime)s - %(levelname)s - %(message)s',  # Log format
        datefmt='%Y-%m-%d %H:%M:%S'  # Date format
    )

    # TODO: implement script
    print('Running')
    time.sleep(5)

    return 0