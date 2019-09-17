# Class that saves folder name and it's path as it's variables
# No struct(exists in C++) in python to simplify this

class folder:
    def __init__(self, name, path):
        self.name = name
        self.path = path

