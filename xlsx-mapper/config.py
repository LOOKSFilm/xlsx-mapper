import os
import getpass
from EditShareAPI import EsAuth
import pickle

osusername = os.getlogin()

def configure():
    if not os.path.exists(os.path.dirname(f"C:/Users/{osusername}/.xlsx-mapper/")):
        os.makedirs(os.path.dirname(f"C:/Users/{osusername}/.xlsx-mapper/"))
    with open(f"C:/Users/{osusername}/.xlsx-mapper/config.p", "wb") as f:
        username = input("Username: ")
        password = getpass.getpass("Password: ")
        #username, password = args.config
        plist = [username, password]
        connection = EsAuth.login("10.0.77.14", username, password)
        if connection == 200:
            print("Logged in on ES-Server")
            pickle.dump(plist, f)
        else:
            print("Wrong username or password")
            sys.exit()


