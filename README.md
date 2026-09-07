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

**1. Delete enrollments older than 3 days & send link automatically** - Add this code in the Apps Script provided in Google Sheets 
```javascript
function deleteOldRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const dateColumn = 1;
  const today = new Date();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    let cellDate = new Date(data[i][dateColumn - 1]);
    if (!isNaN(cellDate) && (today - cellDate) / (1000 * 60 * 60 * 24) > 3) {
      sheet.deleteRow(i + 1);
    }
  }
}
function sendCourseLinks(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  let email, coursesRaw;
  if (e && e.values) {
    email = e.values[1];
    coursesRaw = e.values[2];
  } else {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    email = sheet.getRange(lastRow, 2).getValue();
    coursesRaw = sheet.getRange(lastRow, 3).getValue();
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
  }
}
```
**2. Auto-Sort new Courses** - Add this code in the Apps Script provided in Google Form creation page
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
