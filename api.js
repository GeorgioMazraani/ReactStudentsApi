const express = require("express");
const fs = require("fs");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
app.use(cors());
const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, "students.xlsx");

app.use(express.json());

function readStudents() {
  if (!fs.existsSync(DATA_FILE)) return [];

  const workbook = XLSX.readFile(DATA_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet);
}

function writeStudents(data) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Students");
  XLSX.writeFile(workbook, DATA_FILE);
}

app.get("/students", (req, res) => {
  try {
    const students = readStudents();
    res.json({ success: true, data: students });
  } catch (error) {
    res.status(500).json({ success: false, message: "Error reading students." });
  }
});

app.post("/students", (req, res) => {
  try {
    const { id, name, gender } = req.body;
    if (!id || !name || !gender) {
      return res.status(400).json({ success: false, message: "All fields are required: id, name, gender." });
    }

    let students = readStudents();
    const studentExists = students.some(student => student.id === Number(id));
    if (studentExists) {
      return res.status(400).json({ success: false, message: "Student ID already exists." });
    }

    students.push({ id: Number(id), name, gender }); 
    writeStudents(students);

    res.status(201).json({ success: true, message: "Student added successfully." });
  } catch (error) {
    res.status(500).json({ success: false, message: "Error writing student data." });
  }
});

app.delete("/students/:id", (req, res) => {
  try {
    const studentId = Number(req.params.id);
    let students = readStudents();
    const initialLength = students.length;

    students = students.filter(student => student.id !== studentId);

    if (students.length === initialLength) {
      return res.status(404).json({ success: false, message: "Student not found." });
    }

    writeStudents(students);
    res.json({ success: true, message: "Student removed successfully." });
  } catch (error) {
    res.status(500).json({ success: false, message: "Error deleting student." });
  }
});

app.put("/students/:id", (req, res) => {
  try {
    const studentId = Number(req.params.id);
    const { name, gender } = req.body;

    let students = readStudents();
    const studentIndex = students.findIndex(student => student.id === studentId);

    if (studentIndex === -1) {
      return res.status(404).json({ success: false, message: "Student not found." });
    }

    if (name) students[studentIndex].name = name;
    if (gender) students[studentIndex].gender = gender;

    writeStudents(students);

    res.json({ success: true, message: "Student updated successfully.", data: students[studentIndex] });
  } catch (error) {
    res.status(500).json({ success: false, message: "Error updating student." });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
