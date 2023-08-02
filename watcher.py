import time
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class MyHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        if event.is_directory:
            return
        elif event.event_type == 'created':
            print(f"New file created: {event.src_path}")
            convert_to_pdf(event.src_path)

def convert_to_pdf(file_path):
    root_folder = r"C:\Users\dylan\OneDrive\Desktop\test\watchfolder"
    extension = ".docx"
    subprocess.run(['python', 'docx2pdf.py', root_folder, extension])

# Set the path of the folder to watch
folder_path = r"C:\Users\dylan\OneDrive\Desktop\test\watchfolder"

# Create an observer and attach the event handler
event_handler = MyHandler()
observer = Observer()
observer.schedule(event_handler, folder_path, recursive=False)

# Start the observer to begin monitoring the folder
observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer.join()
