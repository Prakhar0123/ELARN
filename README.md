This project is based on the topic-Online Course Reservation that follows an educational approach to allow students to enroll & learn online,thus reducing the traditional form-filling system that costs both energy & time.New updates will be included further in the project.
___

## 🚀 Getting Started

Clone the repository and install dependencies:

```bash
# Clone the repo
git clone https://github.com/Prakhar0123/ELARN.git

# Navigate into the project folder
cd ELARN

# Install dependencies
npm install

# Start the development server
npm start
```
## Google Apps Script Integration:

I used Google Apps Script to automate tasks in ELARN’s Google Sheets' & Google Forms' backend. Here are the script snippets:

**1. Delete enrollments older than 3 days & send link automatically** - Add & Save this code in the Apps Script provided in Google Sheets > Open Triggers: choose "sendCourseLinks" function, event source "From Spreadsheet", event type "On Form Submit" & Save > Open Editor & Run
```javascript
function sendCourseLinks(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  let email, coursesRaw, rowToDelete;
  if (e && e.range) {
    rowToDelete = e.range.getRow();
    email = e.values[1];
    coursesRaw = e.values[2];
  } else {
    rowToDelete = sheet.getLastRow();
    if (rowToDelete < 2) return;
    email = sheet.getRange(rowToDelete, 2).getValue();
    coursesRaw = sheet.getRange(rowToDelete, 3).getValue();
  }
  if (!email || !coursesRaw) return;
  const maxRow = sheet.getLastRow();
  if (maxRow < 2) return;
  const courseData = sheet.getRange(2, 5, maxRow - 1, 2).getValues();
  const courseMap = {};
  courseData.forEach(row => {
    if (row[0]) {
      courseMap[row[0].toString().trim()] = row[1];
    }
  });
  const enrolledCourses = coursesRaw.toString().split(',').map(c => c.trim());
  let message = "Hello Learner,Greetings from ELARN\n\nHere are the access links for your preferred courses:\n\n";
  let validCoursesFound = false;
  enrolledCourses.forEach(course => {
    if (courseMap[course]) {
      message += `• ${course}: ${courseMap[course]}\n`;
      validCoursesFound = true;
    }
  });
  if (validCoursesFound) {
    MailApp.sendEmail({
      to: email,
      subject: "ELARN-Your Enrolled Course Links",
      body: message
    });
        Utilities.sleep(60000);
    if (rowToDelete > 1) {
      sheet.deleteRow(rowToDelete);
    }
  }
}
```
**2. Auto-Sort new Courses** - Add this code in the Apps Script provided in Google Form creation page & run after adding new course(s) in Forms
```javascript
function autoSortCheckbox() {
  try {
    var form = FormApp.openById("Your form ID from the URL located between d/.../edit");
    var items = form.getItems(FormApp.ItemType.CHECKBOX);
    items.forEach(function(item) {
      var checkboxItem = item.asCheckboxItem();
      var choices = checkboxItem.getChoices();
      var choiceValues = choices.map(function(choice) {
        return choice.getValue();
      });
      choiceValues.sort(function(a, b) {
        return a.toLowerCase().localeCompare(b.toLowerCase());
      });
      var newChoices = choiceValues.map(function(value) {
        return checkboxItem.createChoice(value);
      });
      checkboxItem.setChoices(newChoices);
    });
    Logger.log("Checkbox options sorted successfully.");
  } catch (err) {
    Logger.log("Error: " + err.message);
  }
}
```
---
[!USAGE INSTRUCTIONS (ADMIN)]:
1. The form response sheet must be like:
   <img width="1113" height="288" alt="image" src="https://github.com/user-attachments/assets/bf712e06-016a-4351-841c-bb45e243fce2" />
   Name E & F columns & fill the data manually.
2. Add courses via Forms > Open Apps Script & Run code to sort the courses.
3. In case if the sheets fail to auto-delete courses, open extensions > Apps Script > Run code to restart the process
