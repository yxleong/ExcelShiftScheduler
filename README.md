# Excel Shift Scheduler

Generate monthly calendars within a specified range, organize shifts weekly with checkboxes for check-ins, and send automated LINE notification every Saturday to provide task updates using Apps Script.

## Motivation

I created this project to address the need for an efficient way to arrange dorm cleaning schedules without the manual input of dates. The goal is to automate the process by allowing Excel to generate calendars and organize cleaning shifts automatically. This project aims to streamline the scheduling process and save valuable time and effort.

## Features

- **Generate Monthly Calendars**: Automatically creates monthly calendars with scheduled shifts and checkboxes for cleaning tasks.
- **Dynamic Date Coloring**: Updates the current date to red color daily, ensuring easy identification and tracking of the present day in the calendar.
- **Scheduled LINE Notifications**: Sends LINE notifications every Saturday to remind users of their cleaning tasks.
- **Simple Feedback Mechanism**: Promotes effective communication and enables users to share thoughts on cleaning results.

## Demo

![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/14e67f46-f7b3-46d8-9701-3e6930ce6839)

## Getting Started

To begin using this project, you'll need the following:

1. **Excel Sheet:** Prepare an Excel sheet with the necessary data for generating calendars and organizing cleaning shifts.
![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/cf5ec297-cf1f-4584-ae5b-6a0e9de9cccc)
It's recommended to title your data in **red** color and place it at cell **L2** for optimal compatibility. Also, change the sheet name "calendar"
![image](https://github.com/yxleong/ExcelShiftScheduler/assets/95266740/3acb6739-01c8-45ea-a01e-719123f84ec5)
Feedback sheet for your reference.


3. **LINE Notify Access Token:** Obtain an access token for [LINE Notify](https://notify-bot.line.me/en/) to enable automated notifications.
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

## Completed Improvements

As part of our commitment to continuous improvement, the following enhancements have been made to the project:

- **Update Calendar Functionality:** Implemented an update function for the calendars. This function highlights today's date in red and reverts yesterday's date to its original black color, improving efficiency and reducing redundancy. (Implemented on Nov 18, 2023)

- **Personalized Reminders via Line Notify:** We have enhanced our notification system to include personalized reminders. By utilizing Line Notify to mention individuals by name, we ensure that communications are direct and significantly more effective. This feature enhances the user experience by providing targeted reminders for specific tasks. (Implemented on Nov 18, 2023)

- **Feedback Platform for Cleaning Results:** A feedback platform is aimed at fostering an environment where we can openly provide feedback without any hesitation or fear of embarrassment. It not only facilitates valuable insights into how we perceive the cleanliness of our shared spaces but also promotes a culture of continuous improvement among us. By closely monitoring satisfaction levels and identifying areas that require attention, we are better equipped to make adjustments that meet our collective expectations and improve our living situation. (Implemented on Nov 18, 2023)

## Potential Improvements

As we strive to enhance our shared living experience, we acknowledge the ongoing development and refinement of our project. We've identified a critical area for improvement:

- **Feedback Platform Bug Fixes:** Currently, our feedback platform for cleaning results is experiencing technical difficulties, particularly with the feedback reception function. Users may encounter error messages such as "Exception: Service Spreadsheets failed while accessing document," "We're sorry, a server error occurred. Please wait a bit and try again," or "We're sorry, a server error occurred while reading from storage. Error code INTERNAL." These issues prevent the platform from reliably capturing and storing feedback, impacting our ability to review and act on the cleaning results effectively. We are prioritizing the resolution of these bugs to restore full functionality to the platform, ensuring that all roommates can submit their feedback without encountering errors.

We welcome contributions and ideas from the community to make this project even better.

## Contact

For questions, support, or feedback, you can reach us at [yx0829leong@gmail.com](mailto:yx0829leong@gmail.com). Feel free to [open an issue](../../issues) as well.
