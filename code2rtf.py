import ctypes
import tkinter
from functools import partial
from tkinter.filedialog import askopenfilename

import win32clipboard
from pygments import highlight
from pygments.formatters import RtfFormatter
from pygments.lexers import (get_lexer_by_name, get_lexer_for_filename,
                             guess_lexer)

ctypes.windll.shcore.SetProcessDpiAwareness(True)
# Setup
root = tkinter.Tk()
root.geometry("1024x768")
filePath = None
selected_lexer = tkinter.StringVar()
# Used to make title of the application
applicationName = "Copy Text and Highlight"
root.title(applicationName)
# Setting the font and Padding for the Text Area
fontName = "Consolas"
padding = 10
# Infos about the Document are stored here
document = None
CF_RTF = win32clipboard.RegisterClipboardFormat("Rich Text Format")


def select_lexer(lex):
    global selected_lexer
    selected_lexer = lex


def get_lexer():
    global filePath, selected_lexer
    if selected_lexer.get() != "":
        chosen_lexer = get_lexer_by_name(selected_lexer.get())
    elif filePath is not None:
        chosen_lexer = get_lexer_for_filename(filePath)
    else:
        chosen_lexer = guess_lexer(textArea.selection_get())

    changeWindowTitle(chosen_lexer)
    return chosen_lexer


# Copy File Handler Copy File Handlerss
def formatAndCopy(action=None):
    formatter = RtfFormatter(fontface="Consolas", style="xcode")
    lexer = get_lexer()
    selection = textArea.selection_get()
    result = highlight(selection, lexer, formatter)
    rtf = bytearray(result, encoding="utf-8")
    win32clipboard.OpenClipboard(0)
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(CF_RTF, rtf)
    win32clipboard.CloseClipboard()


def changeWindowTitle(lexer):
    root.title(f"{applicationName} - {filePath} - {lexer}")


# Handle File Events
def fileManager(event=None, action=None):
    global document, filePath
    if action == "open":
        # ask the user for a filename with the native file explorer.
        filePath = askopenfilename()
        with open(filePath, "r") as f:
            document = f.read()
        # Delete Content
        textArea.delete("1.0", tkinter.END)
        # Set Content
        textArea.insert("1.0", document)
        # Set Title


textArea = tkinter.Text(root, font=f"{fontName} 12", relief=tkinter.FLAT)
textArea.pack(fill=tkinter.BOTH, expand=tkinter.TRUE, padx=padding, pady=padding)

menu = tkinter.Menu(root)
root.config(menu=menu)
menu_items = tkinter.Menu(menu)
fileMenu = tkinter.Menu(menu, tearoff=0)
menu.add_cascade(label="File", menu=fileMenu)
menu.add_cascade(label="Lexers", menu=menu_items)

menu_items.add_radiobutton(label="Java", variable=selected_lexer, value="Java")
menu_items.add_radiobutton(label="Rust", variable=selected_lexer, value="Rust")
menu_items.add_radiobutton(label="C++", variable=selected_lexer, value="C++")
menu_items.add_radiobutton(label="Python", variable=selected_lexer, value="Python")
menu_items.add_radiobutton(label="XML", variable=selected_lexer, value="XML")
menu_items.add_radiobutton(label="HTML", variable=selected_lexer, value="HTML")
menu_items.add_radiobutton(label="PHP", variable=selected_lexer, value="PHP")
menu_items.add_radiobutton(label="SQL", variable=selected_lexer, value="SQL")
menu_items.add_radiobutton(
    label="Javascript", variable=selected_lexer, value="Javascript"
)
menu_items.add_radiobutton(label="YAML", variable=selected_lexer, value="YAML")


fileMenu.add_command(
    label="Open", command=partial(fileManager, action="open"), accelerator="Ctrl+O"
)
root.bind_all("<Control-o>", partial(fileManager, action="open"))

fileMenu.add_command(
    label="Copy Selection",
    command=partial(formatAndCopy, action="selection"),
    accelerator="Ctrl+C",
)
root.bind_all("<Control-c>", partial(formatAndCopy))

fileMenu.add_command(label="Exit", command=root.quit)
formatMenu = tkinter.Menu(menu, tearoff=0)

root.mainloop()
