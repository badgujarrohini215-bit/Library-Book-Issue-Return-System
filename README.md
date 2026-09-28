# 📚 Library Book Issue & Return System

A simple **Library Book Issue & Return Management System** developed using **Python, Tkinter, OpenPyXL, and tkcalendar**.

This mini project provides a graphical user interface (GUI) to manage library book records. It allows the user to issue books, return books, view all records, update return details, and delete records. All records are stored in an Excel file for easy management.

---

## 🚀 Features

* 🔐 Login system with username and password
* 📖 Add and issue a new book
* 🔄 Return an existing book
* ✏️ Update/return book records
* 🗑️ Delete book records with confirmation
* 📋 View all library records
* 📅 Select Issue Date and Return Date using a calendar
* 📊 Store records in an Excel file
* ❌ Duplicate Book ID validation
* ⚠️ Missing data validation
* 💬 Success, warning, and error message boxes
* 🔓 Logout functionality
* 🎨 Simple and user-friendly light-colored GUI

---

## 🛠️ Technologies Used

| Technology | Purpose                        |
| ---------- | ------------------------------ |
| Python     | Main programming language      |
| Tkinter    | GUI development                |
| OpenPyXL   | Reading and writing Excel data |
| tkcalendar | Date selection                 |
| Excel      | Data storage                   |
| OS Module  | Checking Excel file existence  |

---

## 📂 Project Structure

```text
Library-Book-Issue-Return-System/
│
├── Library Issue_return.py
├── library_records.xlsx
└── README.md
```

### Files Description

**Library Issue_return.py**
Contains the complete Python application including login, GUI, book issue, return, update, delete, and record display functions.

**library_records.xlsx**
Stores all library book records.

**README.md**
Contains project information, features, technologies, installation steps, and usage instructions.

---

## 📊 Excel Data Structure

The application automatically creates an Excel file named:

```text
library_records.xlsx
```

The following fields are stored:

```text
Book ID
Book Name
Book Author
Student/Person Name
Issue Date
Return Date
Status
```

The Excel file is created automatically if it does not already exist.

---

## 🖥️ Main Modules

### 1. Login

The system starts with a login page.

Default login credentials:

```text
Username: admin
Password: admin123
```

The login system checks the entered username and password before opening the dashboard.

---

### 2. Add / Issue Book

The **Add Record / Issue Book** option opens the book details form.

The user enters:

* Book ID
* Book Name
* Book Author
* Student/Person Name
* Issue/Return option
* Issue Date or Return Date

The application checks whether the Book ID already exists before adding a new record.

---

### 3. Date Selection

The project uses **tkcalendar DateEntry** for selecting dates.

Date format:

```text
DD-MM-YYYY
```

The calendar is available for:

* Issue Date
* Return Date

This avoids manually entering dates and helps maintain a consistent date format.

---

### 4. Update / Return Book

The **Update / Return Book** option displays existing records.

After selecting a record, the book details are loaded into the form. The user can then select a Return Date and return the book.

The status is changed to:

```text
Returned
```

The return date is also stored in the Excel file.

---

### 5. Delete Record

The **Delete Record** option displays all available records.

When the user selects a record, a confirmation popup appears:

```text
Are you sure you want to delete Book ID?
```

The record is deleted only after confirmation.

---

### 6. Show All Records

The **Show All Records** option displays library records in a Tkinter Treeview table.

The table contains:

```text
Book ID
Book Name
Book Author
Student/Person
Issue Date
Return Date
Status
```

The table also includes a scrollbar for viewing records.

---

## 🔄 Project Workflow

```text
                START
                  │
                  ▼
              Login Page
                  │
                  ▼
          Username & Password
                  │
          ┌───────┴────────┐
          │                │
       Valid            Invalid
          │                │
          ▼                ▼
      Dashboard       Error Message
          │
          ▼
    ┌─────┼──────┬──────────┬─────────────┐
    │     │      │          │             │
   Add  Update  Delete    Show All      Clear
    │   /Return  Record    Records       Reset
    │     │      │          │
    ▼     ▼      ▼          ▼
  Issue  Return Delete    Display
  Book   Book   Record     Records
    │     │      │          │
    └─────┴──────┴──────────┘
                  │
                  ▼
            Excel File
                  │
                  ▼
                 END
```

---

## 🧩 Important Functions

The project is divided into separate functions to make the program easier to understand and manage.

| Function                | Purpose                             |
| ----------------------- | ----------------------------------- |
| `create_excel_file()`   | Creates the Excel file and headings |
| `clear_form()`          | Clears form fields                  |
| `clear_screen()`        | Clears the current screen           |
| `show_date_field()`     | Displays the required date field    |
| `show_add_form()`       | Opens Add Record section            |
| `show_update_form()`    | Opens Update/Return section         |
| `show_delete_records()` | Opens Delete Record section         |
| `show_all_records()`    | Displays all records                |
| `perform_action()`      | Handles issue and return operations |
| `load_records()`        | Loads Excel records into the table  |
| `select_record()`       | Handles selected records            |
| `delete_record()`       | Deletes selected records            |
| `logout()`              | Logs the user out                   |
| `login()`               | Validates login credentials         |

---
## ⚙️ Installation : Install Required Libraries

Installed `openpyxl`:

Installed `tkcalendar`

Tkinter generally comes with the standard Python installation on Windows.

---

* Login using:

```text
Username: admin
Password: admin123
```

* Use the dashboard options to manage library records.

---

## 💾 Data Storage

The application uses an Excel workbook for data storage instead of a database.

The Excel file is automatically created when the program is started if it does not already exist.

All records are stored in:

```text
library_records.xlsx
```

---

## 🎯 Project Objectives

The main objectives of this project are:

* To develop a simple library management application.
* To understand GUI development using Tkinter.
* To understand file handling using OpenPyXL.
* To use date selection through a calendar widget.
* To create a user-friendly interface for managing book records.

---

**Library Book Issue & Return System**

Developed as a Python mini project to demonstrate GUI development, Excel file handling, and basic library record management.

---

## ⭐ Project Highlights

> A simple, practical, and user-friendly library management application built using Python GUI and Excel-based data storage.
