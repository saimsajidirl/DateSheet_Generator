# 🗓️ Exam Date Sheet Generator

This Python script allows you to easily create a professional-looking exam date sheet in Microsoft Word format (`.docx`). Using the `python-docx` library, the script interactively collects exam details and automatically formats them into a structured table.

## 📋 Features

- Creates a `.docx` file with the heading **Exam Date Sheet**
- Dynamically collects data for any number of exams
- Adds a table with the following columns:
  - Date
  - Time
  - Course
  - Session
  - Room
  - Teacher Name
- Saves the output as `Date_Sheet.docx`

---

## 🚀 Getting Started

### Prerequisites

Make sure you have Python installed. Then, install the required library:

```bash
pip install python-docx
```

---

### 🧠 How It Works

1. Run the script.
2. Enter the number of exams.
3. For each exam, enter:
   - Date (e.g., `25-Nov-2024`)
   - Time (e.g., `10:00 AM`)
   - Course name
   - Semester & Section (e.g., `Fall 2024, Section A`)
   - Room number
   - Teacher name
4. A `.docx` file named `Date_Sheet.docx` is created in the same directory.

---

## 📎 Example Usage

```bash
python exam_date_sheet.py
```

Example Input:
```
Enter the number of exams: 2

Enter details for exam 1:
Enter the date (e.g., 25-Nov-2024): 25-Nov-2024
Enter the time (e.g., 10:00 AM): 10:00 AM
Enter the course name: Physics
Enter the semester & section (e.g., Fall 2024, Section A): Fall 2024, Section A
Enter the room number: B-203
Enter the teacher's name: Mr. Usman

Enter details for exam 2:
Enter the date (e.g., 25-Nov-2024): 27-Nov-2024
Enter the time (e.g., 10:00 AM): 1:00 PM
Enter the course name: Mathematics
Enter the semester & section (e.g., Fall 2024, Section A): Fall 2024, Section A
Enter the room number: B-204
Enter the teacher's name: Ms. Ayesha
```

Output:
A Word document named `Date_Sheet.docx` with a formatted table.

---

## 📄 Sample Output

| Date         | Time     | Course     | Session               | Room  | Teacher Name |
|--------------|----------|------------|------------------------|-------|----------------|
| 25-Nov-2024  | 10:00 AM | Physics    | Fall 2024, Section A   | B-203 | Mr. Usman      |
| 27-Nov-2024  | 1:00 PM  | Mathematics| Fall 2024, Section A   | B-204 | Ms. Ayesha     |

---

## 🤝 Contributing

Feel free to fork the repository and submit a pull request for improvements or new features!

---

## 📜 License

This project is open source and available under the [MIT License](LICENSE).