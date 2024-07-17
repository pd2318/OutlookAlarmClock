import win32com.client
import schedule
import time
import threading
import tkinter as tk
from tkinter import messagebox, Listbox, Scrollbar, Entry, Button, Label
import pygame
from datetime import datetime, timedelta

# Initialize pygame mixer
pygame.mixer.init()

# Path to your alarm sound
ALARM_SOUND_PATH = 'C:/Users/pdres/Music/alarm_sound.wav'

# Global variables for GUI elements
alarm_listbox = None
custom_meeting_entry = None
custom_time_entry = None
no_meetings_label = None

# Flag to control alarm state
alarm_playing = False

def get_outlook_meetings():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)  # 9 refers to the calendar folder

    today = datetime.now().date()
    begin = today.strftime("%m/%d/%Y")
    tomorrow = (today + timedelta(days=1)).strftime("%m/%d/%Y")
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    # Restrict the items to today's meetings
    restriction = "[Start] >= '" + begin + "' AND [End] <= '" + tomorrow + "'"
    items = items.Restrict(restriction)

    meetings = []
    for item in items:
        if item.Start.date() == today:
            meetings.append(item)

    return meetings

def refresh_alarm_list():
    global alarm_listbox
    alarm_listbox.delete(0, tk.END)
    jobs = schedule.get_jobs()
    for job in jobs:
        alarm_listbox.insert(tk.END, f"{job.job_func.args[0]} at {job.job_func.args[1]}")

def main():
    global alarm_listbox, custom_meeting_entry, custom_time_entry, no_meetings_label

    # GUI setup
    root = tk.Tk()
    root.title("Outlook Alarm Clock")

    main_frame = tk.Frame(root)
    main_frame.pack(pady=10)

    scrollbar = Scrollbar(main_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    alarm_listbox = Listbox(main_frame, yscrollcommand=scrollbar.set, width=50, height=10)
    alarm_listbox.pack(side=tk.LEFT)
    scrollbar.config(command=alarm_listbox.yview)

    custom_meeting_entry = Entry(root, width=30)
    custom_meeting_entry.pack(pady=5)
    custom_meeting_entry.insert(0, "Custom Meeting Title")

    custom_time_entry = Entry(root, width=15)
    custom_time_entry.pack(pady=5)
    custom_time_entry.insert(0, "HH:MM")

    add_button = Button(root, text="Add Custom Alarm", command=add_custom_alarm)
    add_button.pack(pady=5)

    delete_button = Button(root, text="Delete Selected Alarm", command=delete_alarm)
    delete_button.pack(pady=5)

    no_meetings_label = Label(root, text="No Meetings Scheduled")
    no_meetings_label.pack()

    # Fetch Outlook meetings and set alarms
    meetings = get_outlook_meetings()
    if meetings:
        no_meetings_label.pack_forget()
        for meeting in meetings:
            set_alarm(meeting.Subject, meeting.Start.strftime("%H:%M"))
            alarm_listbox.insert(tk.END, f"{meeting.Subject} at {meeting.Start.strftime('%H:%M')}")
    else:
        no_meetings_label.pack()

    def run_schedule():
        while True:
            schedule.run_pending()
            time.sleep(1)

    schedule_thread = threading.Thread(target=run_schedule)
    schedule_thread.start()

    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("Exiting...")
        exit(0)

def play_sound():
    global alarm_playing
    pygame.mixer.music.load(ALARM_SOUND_PATH)
    pygame.mixer.music.play(loops=-1)  # Play in a loop until stopped
    alarm_playing = True

def stop_sound():
    global alarm_playing
    pygame.mixer.music.stop()
    alarm_playing = False

def alarm_action(meeting_subject, alarm_time):
    def show_alarm():
        def snooze():
            stop_sound()
            snooze_time = datetime.now() + timedelta(minutes=5)
            set_alarm(meeting_subject, snooze_time.strftime("%H:%M"))
            messagebox.showinfo("Snooze", f"Alarm snoozed for {meeting_subject} at {snooze_time.strftime('%H:%M')}")
            root.destroy()

        def dismiss():
            stop_sound()
            root.destroy()

        root = tk.Tk()
        root.title("Alarm")

        label = tk.Label(root, text=f"Meeting '{meeting_subject}' is starting now!")
        label.pack(pady=10)
        snooze_button = tk.Button(root, text="Snooze", command=snooze)
        snooze_button.pack(side="left", padx=10)
        dismiss_button = tk.Button(root, text="Dismiss", command=dismiss)
        dismiss_button.pack(side="right", padx=10)
        root.protocol("WM_DELETE_WINDOW", dismiss)  # Handle window close event

        # Play sound in a separate thread
        play_thread = threading.Thread(target=play_sound)
        play_thread.start()

        root.mainloop()

    thread = threading.Thread(target=show_alarm)
    thread.start()

def set_alarm(meeting_subject, alarm_time):
    schedule.every().day.at(alarm_time).do(alarm_action, meeting_subject, alarm_time)

def add_custom_alarm():
    global custom_meeting_entry, custom_time_entry
    meeting_subject = custom_meeting_entry.get()
    alarm_time = custom_time_entry.get()
    if meeting_subject and alarm_time:
        set_alarm(meeting_subject, alarm_time)
        custom_meeting_entry.delete(0, tk.END)
        custom_time_entry.delete(0, tk.END)
        alarm_listbox.insert(tk.END, f"{meeting_subject} at {alarm_time}")  # Insert once per addition
        refresh_alarm_list()  # Ensure the list is refreshed
        
def delete_alarm():
    global alarm_listbox
    selected_alarm = alarm_listbox.curselection()
    if selected_alarm:
        alarm_listbox.delete(selected_alarm[0])
        # Logic to remove the alarm from the schedule
        # This is more complex as schedule doesn't have a direct remove function
        # You'd need to keep track of jobs and manually remove the correct one

if __name__ == "__main__":
    main()
