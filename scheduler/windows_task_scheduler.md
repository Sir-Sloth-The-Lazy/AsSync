# Daily run on Windows

1. Open **Task Scheduler** → *Create Task* (not "Basic Task").
2. **General**
   - Name: `RISE Asana Sync`
   - Select *Run whether user is logged on or not*
   - Tick *Run with highest privileges* only if your Python install needs it
3. **Triggers** → New
   - Daily, start `07:30`, recur every 1 day
   - Tick *Stop task if it runs longer than* `30 minutes`
4. **Actions** → New
   - Action: *Start a program*
   - Program/script: `C:\path\to\rise-asana-sync\scheduler\run_sync.bat`
   - Start in: `C:\path\to\rise-asana-sync`
   (Set "Start in" — without it the script cannot find config.yaml.)
5. **Conditions**
   - Untick *Start the task only if the computer is on AC power* if you use a laptop
   - Tick *Wake the computer to run this task* if you want it to run while asleep
6. **Settings**
   - Tick *Run task as soon as possible after a scheduled start is missed* — this
     is what covers days your laptop was closed at 07:30

Check it worked: `logs\sync.log` and `build\last_run.json`.
