import window
import tkinter as tk

# TODO: Restructure Geometry

root = tk.Tk()  # Main window
my_gui = window.MainWindow(root)
my_gui.pack(side="top", fill="both", expand=True)

# Gets the requested values of the height and width.
windowWidth = root.winfo_reqwidth()
windowHeight = root.winfo_reqheight()
print("Width", windowWidth, "Height", windowHeight)

# Gets both half the screen width/height and window width/height
positionRight = int(root.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(root.winfo_screenheight() / 2 - windowHeight / 2)

# Positions the window in the center of the page.
root.geometry("+{}+{}".format(positionRight-150, positionDown-100))
# Hold window open until we close it
root.mainloop()
