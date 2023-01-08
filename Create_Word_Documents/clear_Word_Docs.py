import os
from shutil import rmtree

def main():
    location = os.getcwd()
    dir = "Word_Docs"
    path = os.path.join(location, dir)

    remove_directory_contents(path)

def remove_directory_contents(path: str) -> None:
    """
    A function meant to clear the contents of a directory. It will not delete
    the directory itself.
    """
    if not isinstance(path, str): raise TypeError("Must provide a valid path")
    if not os.path.isdir(path): raise TypeError("Must provide a directory")

    for filename in os.listdir(path):
        filepath = os.path.join(path, filename)
        # If the filepath leads to a file
        if os.path.isfile(filepath):
            os.unlink(filepath)

        # If the filepath leads to a directory
        if os.path.isdir(filepath):
            rmtree(filepath)

if __name__ == '__main__':
    main()