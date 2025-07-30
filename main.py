from late import Late
from late_v2 import Additional_Late
from rush import Rush
import time
from worklist import Worklist

def main():
    Worklist()
    time.sleep(2)
    Rush()
    time.sleep(2)
    Late()
    time.sleep(2)
    Additional_Late()


if __name__ == "__main__":
    main()
