# Excel Shift Scheduler

Generate monthly calendars within a specified range, organize shifts weekly with checkboxes for check-ins, and send automated LINE notification every Saturday to provide task updates using Apps Script.

## Motivation

I created this project to address the need for an efficient way to arrange dorm cleaning schedules without the manual input of dates. The goal is to automate the process by allowing Excel to generate calendars and organize cleaning shifts automatically. This project aims to streamline the scheduling process and save valuable time and effort.

## Features

- Generate monthly calendars with shifts scheduled and checkboxes.
- Automatically update the current date to red color daily.
- Send LINE notification every Saturday.

## Demo

![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/14e67f46-f7b3-46d8-9701-3e6930ce6839)

## Getting Started

To begin using this project, you'll need the following:

1. **Excel Sheet:** Prepare an Excel sheet with the necessary data for generating calendars and organizing cleaning shifts.
![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/cf5ec297-cf1f-4584-ae5b-6a0e9de9cccc)
It's recommended to title your data in **red** color and place it at cell **L2** for optimal compatibility.

2. **LINE Notify Access Token:** Obtain an access token for [LINE Notify](https://notify-bot.line.me/en/) to enable automated notifications.
![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/5505319b-90c2-4d33-bac5-81317d91b5b8)
![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/aa123cb2-0036-4a89-97e8-8711c8abed27)

With these prerequisites in place, you can easily get started with the project.

## Usage

Follow these steps to use this project in Google Sheets with Apps Script:

1. **Open Apps Script Editor:**
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/667abe1f-efe4-48f4-86c4-aa4f167b9c4d)
   - Click on `Extensions` in the top menu.
   - Select `Apps Script` to open the script editor.

2. **Create and Paste the Code:**
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/0d1e1d94-1d37-4822-a6a4-845120a8c478)
   - Create two script files, `calendar.gs` and `line.gs`, if you haven't already.
   - Paste the respective code into each file.

3. **Run `calendar.gs`:**
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/1f98ec9c-8030-4c15-ad4f-b890322d65a4)
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/40891fd1-e360-45da-b0a8-d21fb7c70496)
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/39d704d4-0a6c-4615-b847-8802d5d0a171)
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/431ae8e1-2c8f-4765-916f-2e32f9275173)
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/3264371e-3986-4b49-8748-fca70d31c739)
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/5b0b9aa1-3966-4005-86cb-f438efc4adbd)
   - In the Apps Script editor, open `calendar.gs`.
   - Click the save (💾) button to save the project.
   - Review permissions: 
     - Click `Advanced` in the permissions dialog.
     - Go to `Untitled project (unsafe)`.
     - Allow permissions.
   - Click the run (▶️) button to execute the script and generate the calendar.
   - Wait for the process to complete.

4. **Set Up Triggers:**
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/03044d72-7209-4a61-8602-d4112c67d4ce)
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/0749e6fa-328b-4503-ab63-838a8492819d)
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/44bb5e00-8273-4c47-94a0-b41ab7140716)
  ![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/07ed9cea-d8b6-4110-ae8a-844737287fc5)
   - In the Apps Script editor, navigate to `Triggers`.
   - Add triggers as needed to automate the scheduling:
     - Choose which function to run: `generateCalendar`.
     - Select event source: `Time-driven`.
     - Select type of time-based trigger: `Day timer`.
     - Select time of day: `Midnight to 1am`.
   - Add triggers as needed to automate the LINE notifications:
     - Choose which function to run: `checkCheckboxStatus`.
     - Select event source: `Time-driven`.
     - Select type of time-based trigger: `Week timer`.
     - Select day of the week: `Every Saturday`.

## Potential Improvements

While the project is functional, there are areas and features where improvements can be made.

Some potential enhancements include:

- **Update Calendar Functionality:** Instead of generating new calendars every time the code runs, consider implementing a function to update the existing calendars. This function could highlight today's date in red and revert yesterday's date to its original black color. This approach would improve efficiency and reduce redundancy.

We welcome contributions and ideas from the community to make this project even better.

## Contact

For questions, support, or feedback, you can reach us at [yx0829leong@gmail.com](mailto:yx0829leong@gmail.com). Feel free to [open an issue](../../issues) as well.
