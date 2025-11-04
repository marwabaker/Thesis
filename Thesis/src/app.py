from __future__ import annotations

import csv
import re
import sys
from dataclasses import dataclass
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from PyQt5 import QtCore, QtWidgets

from .ui_mainwindow import Ui_MainWindow


@dataclass
class Professor:
    prof_id: str
    name: str
    office_start: time
    office_end: time

    def as_row(self) -> List[str]:
        return [
            self.prof_id,
            self.name,
            f"{self.office_start.strftime('%H:%M')} - {self.office_end.strftime('%H:%M')}",
        ]

    def as_dict(self) -> dict:
        office_hours = f"{self.office_start.strftime('%H:%M')} - {self.office_end.strftime('%H:%M')}"
        return {
            "PROF_ID": self.prof_id,
            "PROF_NAME": self.name,
            "OFFICE_START": self.office_start.strftime("%H:%M"),
            "OFFICE_END": self.office_end.strftime("%H:%M"),
            "OFFICE_HOURS": office_hours,
        }


@dataclass
class Room:
    number: str
    capacity: int
    lec_sem: str

    def as_row(self) -> List[str]:
        return [self.number, str(self.capacity), self.lec_sem]

    def as_dict(self) -> dict:
        return {
            "ROOM_NUMBER": self.number,
            "CAPACITY": str(self.capacity),
            "LEC_SEM": self.lec_sem,
        }


@dataclass
class Subject:
    code: str
    name: str
    subject_type: str
    semester: int
    format_: str
    capacity: int
    professor_id: str

    def as_row(self) -> List[str]:
        return [
            self.code,
            self.name,
            self.subject_type,
            str(self.semester),
            self.format_,
            str(self.capacity),
            self.professor_id,
        ]

    def as_dict(self) -> dict:
        return {
            "SUBJECT_CODE": self.code,
            "SUBJECT_NAME": self.name,
            "TYPE": self.subject_type,
            "SEMESTER": str(self.semester),
            "LEC_SEM": self.format_,
            "CLASS_CAP": str(self.capacity),
            "PROF_ID": self.professor_id,
        }


@dataclass
class ScheduleEntry:
    schedule_id: str
    subject_code: str
    professor_id: str
    room_number: str
    day: str
    time: str

    def as_row(self) -> List[str]:
        return [
            self.schedule_id,
            self.subject_code,
            self.professor_id,
            self.room_number,
            self.day,
            self.time,
        ]

    def as_dict(self) -> dict:
        return {
            "SCHEDULE_ID": self.schedule_id,
            "SUBJECT_CODE": self.subject_code,
            "PROF_ID": self.professor_id,
            "ROOM_NUMBER": self.room_number,
            "DAY": self.day,
            "TIME": self.time,
        }


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.project_dir = Path(__file__).resolve().parent
        self.data_dir = (self.project_dir.parent / "data").resolve()
        self.data_dir.mkdir(exist_ok=True)
        self.professors_path = self.data_dir / "professors.csv"
        self.rooms_path = self.data_dir / "rooms.csv"
        self.subjects_path = self.data_dir / "subjects.csv"
        self.schedule_path = self.data_dir / "schedule.csv"

        self.professors: List[Professor] = []
        self.rooms: List[Room] = []
        self.subjects: List[Subject] = []
        self.schedule_entries: List[ScheduleEntry] = []

        self._configure_tables()
        self._connect_signals()
        self._init_theme_controls()
        self._load_all_data()
        self.statusBar().showMessage("Ready", 3000)

    def _configure_tables(self) -> None:
        tables = (
            self.ui.tableWidget_prof,
            self.ui.tableWidget_rooms,
            self.ui.tableWidget_subjects,
            self.ui.tableWidget_schedule,
        )
        for table in tables:
            table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
            table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
            table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
            table.horizontalHeader().setStretchLastSection(True)
            table.verticalHeader().setVisible(False)
            table.setAlternatingRowColors(True)

    def _connect_signals(self) -> None:
        self.ui.pushButton_import_prof.clicked.connect(self._import_professors)
        self.ui.pushButton_add_prof.clicked.connect(self._add_professor)
        self.ui.pushButton_edit_prof.clicked.connect(self._edit_professor)
        self.ui.pushButton_delete_prof.clicked.connect(self._delete_professor)
        self.ui.tableWidget_prof.itemSelectionChanged.connect(self._sync_professor_form)

        self.ui.pushButton_import_rooms.clicked.connect(self._import_rooms)
        self.ui.pushButton_add_room.clicked.connect(self._add_room)
        self.ui.pushButton_edit_room.clicked.connect(self._edit_room)
        self.ui.pushButton_delete_room.clicked.connect(self._delete_room)
        self.ui.tableWidget_rooms.itemSelectionChanged.connect(self._sync_room_form)

        self.ui.pushButton_import_subjects.clicked.connect(self._import_subjects)
        self.ui.pushButton_add_subject.clicked.connect(self._add_subject)
        self.ui.pushButton_edit_subject.clicked.connect(self._edit_subject)
        self.ui.pushButton_delete_subject.clicked.connect(self._delete_subject)
        self.ui.tableWidget_subjects.itemSelectionChanged.connect(self._sync_subject_form)

        self.ui.pushButton_generate_schd.clicked.connect(self._generate_schedule)
        self.ui.comboBox_choose_prof.currentIndexChanged.connect(self._refresh_schedule_table)
        self.ui.comboBox_choose_room.currentIndexChanged.connect(self._refresh_schedule_table)

    def _load_all_data(self) -> None:
        self._load_professors()
        self._load_rooms()
        self._load_subjects()
        self._load_schedule()
        self._refresh_professor_table()
        self._refresh_room_table()
        self._refresh_subject_table()
        self._refresh_schedule_filters()
        self._refresh_schedule_table()
        self._update_professor_completer()

    # ----------------------------
    # Professors
    # ----------------------------
    def _load_professors(self) -> None:
        self.professors = []
        if not self.professors_path.exists():
            self._show_missing_file(self.professors_path)
            return
        try:
            with self.professors_path.open("r", newline="", encoding="utf-8-sig") as handle:
                reader = csv.DictReader(handle)
                for row in reader:
                    prof_id = row.get("PROF_ID", "").strip()
                    name = row.get("PROF_NAME", "").strip()
                    start_text = (row.get("OFFICE_START") or "").strip()
                    end_text = (row.get("OFFICE_END") or "").strip()
                    if not start_text and not end_text:
                        legacy_hours = (row.get("OFFICE_HOURS") or "").strip()
                        start_text, end_text = self._split_hours(legacy_hours)
                    office_start = self._parse_time_string(start_text, default=time(9, 0))
                    office_end = self._parse_time_string(end_text, default=time(11, 0))
                    office_start, office_end = self._normalize_time_range(office_start, office_end)
                    if prof_id and name:
                        self.professors.append(Professor(prof_id, name, office_start, office_end))
        except Exception as exc:  # pragma: no cover
            self._show_io_error(self.professors_path, exc)

        self.professors.sort(key=lambda prof: prof.prof_id)

    def _refresh_professor_table(self) -> None:
        headers = ["Professor ID", "Name", "working hours"]
        rows = [prof.as_row() for prof in self.professors]
        self._populate_table(self.ui.tableWidget_prof, rows, headers)

    def _sync_professor_form(self) -> None:
        row = self._selected_row(self.ui.tableWidget_prof)
        if row is None:
            self.ui.lineEdit_prof_name.clear()
            self.ui.timeEdit_office_start.setTime(QtCore.QTime(9, 0))
            self.ui.timeEdit_office_end.setTime(QtCore.QTime(11, 0))
            return
        professor = self.professors[row]
        self.ui.lineEdit_prof_name.setText(professor.name)
        self.ui.timeEdit_office_start.setTime(self._qtime_from_time(professor.office_start))
        self.ui.timeEdit_office_end.setTime(self._qtime_from_time(professor.office_end))

    def _add_professor(self) -> None:
        name = self._clean_text(self.ui.lineEdit_prof_name.text(), default="Dr.John Smith")
        if not name:
            self._show_validation_error("Please enter a professor name.")
            return

        office_start = self._time_from_qtime(self.ui.timeEdit_office_start.time())
        office_end = self._time_from_qtime(self.ui.timeEdit_office_end.time())
        if office_start >= office_end:
            self._show_validation_error("Office end time must be after the start time.")
            return

        suggestion = self._suggest_identifier([prof.prof_id for prof in self.professors], prefix="NEW", digits=3)
        prof_id, ok = QtWidgets.QInputDialog.getText(
            self,
            "Professor ID",
            "Enter unique professor ID:",
            text=suggestion,
        )
        if not ok:
            return
        prof_id = self._clean_text(prof_id).upper()
        if not prof_id:
            self._show_validation_error("Professor ID cannot be empty.")
            return
        if any(prof.prof_id == prof_id for prof in self.professors):
            self._show_validation_error(f"Professor ID '{prof_id}' already exists.")
            return

        self.professors.append(Professor(prof_id, name, office_start, office_end))
        self.professors.sort(key=lambda prof: prof.prof_id)
        self._refresh_professor_table()
        self._select_table_row(self.ui.tableWidget_prof, self.professors, "prof_id", prof_id)
        self._save_professors()
        self._refresh_schedule_filters()
        self._update_professor_completer()
        self.statusBar().showMessage(f"Professor {prof_id} added.", 4000)

    def _edit_professor(self) -> None:
        row = self._selected_row(self.ui.tableWidget_prof)
        if row is None:
            self._show_validation_error("Select a professor to edit.")
            return
        professor = self.professors[row]
        name = self._clean_text(self.ui.lineEdit_prof_name.text()) or professor.name
        office_start = self._time_from_qtime(self.ui.timeEdit_office_start.time())
        office_end = self._time_from_qtime(self.ui.timeEdit_office_end.time())
        if office_start >= office_end:
            self._show_validation_error("Office end time must be after the start time.")
            return
        professor.name = name
        professor.office_start = office_start
        professor.office_end = office_end
        self._refresh_professor_table()
        self._select_table_row(self.ui.tableWidget_prof, self.professors, "prof_id", professor.prof_id)
        self._save_professors()
        self._refresh_schedule_filters()
        self._update_professor_completer()
        self.statusBar().showMessage(f"Professor {professor.prof_id} updated.", 4000)

    def _delete_professor(self) -> None:
        row = self._selected_row(self.ui.tableWidget_prof)
        if row is None:
            self._show_validation_error("Select a professor to delete.")
            return
        professor = self.professors[row]
        dependent_subjects = sum(1 for subj in self.subjects if subj.professor_id == professor.prof_id)
        dependent_schedule = sum(1 for entry in self.schedule_entries if entry.professor_id == professor.prof_id)
        warning = (
            f"Delete professor {professor.prof_id}?\n"
            f"Referenced by {dependent_subjects} subject(s) and {dependent_schedule} schedule entry(ies)."
        )
        reply = QtWidgets.QMessageBox.question(self, "Confirm Deletion", warning)
        if reply != QtWidgets.QMessageBox.Yes:
            return
        del self.professors[row]
        self._refresh_professor_table()
        self._save_professors()
        self._refresh_schedule_filters()
        self._update_professor_completer()
        self.statusBar().showMessage("Professor removed.", 4000)

    def _save_professors(self) -> None:
        fieldnames = ["PROF_ID", "PROF_NAME", "OFFICE_START", "OFFICE_END", "OFFICE_HOURS"]
        self._write_csv(self.professors_path, [prof.as_dict() for prof in self.professors], fieldnames=fieldnames)

    def _import_professors(self) -> None:
        path = self._open_file_dialog("Import Professors", self.professors_path)
        if not path:
            return
        self.professors_path = path
        self._load_professors()
        self._refresh_professor_table()
        self._refresh_schedule_filters()
        self._update_professor_completer()
        self.statusBar().showMessage(f"Loaded professors from {path.name}.", 4000)

    # ----------------------------
    # Rooms
    # ----------------------------
    def _load_rooms(self) -> None:
        self.rooms = []
        if not self.rooms_path.exists():
            self._show_missing_file(self.rooms_path)
            return
        try:
            with self.rooms_path.open("r", newline="", encoding="utf-8-sig") as handle:
                reader = csv.DictReader(handle)
                for row in reader:
                    number = row.get("ROOM_NUMBER", "").strip()
                    capacity_text = row.get("CAPACITY", "").strip()
                    lec_sem = row.get("LEC_SEM", "").strip()
                    if not number:
                        continue
                    try:
                        capacity = int(capacity_text)
                    except ValueError:
                        capacity = 0
                    self.rooms.append(Room(number, capacity, lec_sem))
        except Exception as exc:  # pragma: no cover
            self._show_io_error(self.rooms_path, exc)
        self.rooms.sort(key=lambda room: room.number)

    def _refresh_room_table(self) -> None:
        headers = ["Room Number", "Capacity", "Format"]
        rows = [room.as_row() for room in self.rooms]
        self._populate_table(self.ui.tableWidget_rooms, rows, headers)

    def _sync_room_form(self) -> None:
        row = self._selected_row(self.ui.tableWidget_rooms)
        if row is None:
            self.ui.lineEdit_room_number.clear()
            self.ui.lineEdit_room_capacity.clear()
            self.ui.comboBox_room_type.setCurrentIndex(0)
            return
        room = self.rooms[row]
        self.ui.lineEdit_room_number.setText(room.number)
        self.ui.lineEdit_room_capacity.setText(str(room.capacity))
        index = self.ui.comboBox_room_type.findText(room.lec_sem, QtCore.Qt.MatchFixedString)
        if index >= 0:
            self.ui.comboBox_room_type.setCurrentIndex(index)

    def _add_room(self) -> None:
        number = self._clean_text(self.ui.lineEdit_room_number.text(), default="Room Number")
        capacity_text = self._clean_text(self.ui.lineEdit_room_capacity.text(), default="Capacity")
        room_type = self.ui.comboBox_room_type.currentText()
        if not number:
            self._show_validation_error("Please enter a room number.")
            return
        try:
            capacity = int(capacity_text)
        except ValueError:
            self._show_validation_error("Capacity must be a number.")
            return
        if capacity <= 0:
            self._show_validation_error("Capacity must be positive.")
            return
        if room_type not in {"Lecture", "Seminar"}:
            self._show_validation_error("Select Lecture or Seminar for room type.")
            return
        if any(room.number == number for room in self.rooms):
            self._show_validation_error(f"Room {number} already exists.")
            return
        self.rooms.append(Room(number, capacity, room_type))
        self.rooms.sort(key=lambda room: room.number)
        self._refresh_room_table()
        self._select_table_row(self.ui.tableWidget_rooms, self.rooms, "number", number)
        self._save_rooms()
        self._refresh_schedule_filters()
        self.statusBar().showMessage(f"Room {number} added.", 4000)

    def _edit_room(self) -> None:
        row = self._selected_row(self.ui.tableWidget_rooms)
        if row is None:
            self._show_validation_error("Select a room to edit.")
            return
        room = self.rooms[row]
        number = self._clean_text(self.ui.lineEdit_room_number.text()) or room.number
        capacity_text = self._clean_text(self.ui.lineEdit_room_capacity.text()) or str(room.capacity)
        room_type = self.ui.comboBox_room_type.currentText()
        try:
            capacity = int(capacity_text)
        except ValueError:
            self._show_validation_error("Capacity must be a number.")
            return
        if capacity <= 0:
            self._show_validation_error("Capacity must be positive.")
            return
        if room_type not in {"Lecture", "Seminar"}:
            self._show_validation_error("Select Lecture or Seminar for room type.")
            return
        duplicate = next((r for r in self.rooms if r.number == number and r is not room), None)
        if duplicate is not None:
            self._show_validation_error(f"Another room already uses number {number}.")
            return
        room.number = number
        room.capacity = capacity
        room.lec_sem = room_type
        self.rooms.sort(key=lambda r: r.number)
        self._refresh_room_table()
        self._select_table_row(self.ui.tableWidget_rooms, self.rooms, "number", room.number)
        self._save_rooms()
        self._refresh_schedule_filters()
        self.statusBar().showMessage("Room updated.", 4000)

    def _delete_room(self) -> None:
        row = self._selected_row(self.ui.tableWidget_rooms)
        if row is None:
            self._show_validation_error("Select a room to delete.")
            return
        room = self.rooms[row]
        dependent_schedule = sum(1 for entry in self.schedule_entries if entry.room_number == room.number)
        warning = f"Delete room {room.number}?\nReferenced by {dependent_schedule} schedule entry(ies)."
        reply = QtWidgets.QMessageBox.question(self, "Confirm Deletion", warning)
        if reply != QtWidgets.QMessageBox.Yes:
            return
        del self.rooms[row]
        self._refresh_room_table()
        self._save_rooms()
        self._refresh_schedule_filters()
        self.statusBar().showMessage("Room removed.", 4000)

    def _save_rooms(self) -> None:
        self._write_csv(self.rooms_path, [room.as_dict() for room in self.rooms], fieldnames=["ROOM_NUMBER", "CAPACITY", "LEC_SEM"])

    def _import_rooms(self) -> None:
        path = self._open_file_dialog("Import Rooms", self.rooms_path)
        if not path:
            return
        self.rooms_path = path
        self._load_rooms()
        self._refresh_room_table()
        self._refresh_schedule_filters()
        self.statusBar().showMessage(f"Loaded rooms from {path.name}.", 4000)

    # ----------------------------
    # Subjects
    # ----------------------------
    def _load_subjects(self) -> None:
        self.subjects = []
        if not self.subjects_path.exists():
            self._show_missing_file(self.subjects_path)
            return
        try:
            with self.subjects_path.open("r", newline="", encoding="utf-8-sig") as handle:
                reader = csv.DictReader(handle)
                for row in reader:
                    code = row.get("SUBJECT_CODE", "").strip()
                    name = row.get("SUBJECT_NAME", "").strip()
                    subject_type = row.get("TYPE", "").strip()
                    semester_text = row.get("SEMESTER", "0").strip()
                    format_ = row.get("LEC_SEM", "").strip()
                    capacity_text = row.get("CLASS_CAP", "0").strip()
                    professor_id = row.get("PROF_ID", "").strip()
                    if not code:
                        continue
                    try:
                        semester = int(semester_text)
                    except ValueError:
                        semester = 0
                    try:
                        capacity = int(capacity_text)
                    except ValueError:
                        capacity = 0
                    self.subjects.append(
                        Subject(
                            code=code,
                            name=name,
                            subject_type=subject_type,
                            semester=semester,
                            format_=format_,
                            capacity=capacity,
                            professor_id=professor_id,
                        )
                    )
        except Exception as exc:  # pragma: no cover
            self._show_io_error(self.subjects_path, exc)
        self.subjects.sort(key=lambda subject: subject.code)

    def _refresh_subject_table(self) -> None:
        headers = ["Code", "Name", "Type", "Semester", "Format", "Capacity", "Professor ID"]
        rows = [subject.as_row() for subject in self.subjects]
        self._populate_table(self.ui.tableWidget_subjects, rows, headers)

    def _sync_subject_form(self) -> None:
        row = self._selected_row(self.ui.tableWidget_subjects)
        if row is None:
            self.ui.lineEdit_subject_name.clear()
            self.ui.comboBox_subject_type.setCurrentIndex(0)
            self.ui.comboBox_subject_lec_sem.setCurrentIndex(0)
            self.ui.comboBox_subject_semester.setCurrentIndex(0)
            self.ui.lineEdit_subject_cap.clear()
            self.ui.lineEdit_subject_prof_id.clear()
            return
        subject = self.subjects[row]
        self.ui.lineEdit_subject_name.setText(subject.name)
        self._set_combobox_text(self.ui.comboBox_subject_type, subject.subject_type)
        self._set_combobox_text(self.ui.comboBox_subject_lec_sem, subject.format_)
        self._set_combobox_text(self.ui.comboBox_subject_semester, str(subject.semester))
        self.ui.lineEdit_subject_cap.setText(str(subject.capacity))
        self.ui.lineEdit_subject_prof_id.setText(subject.professor_id)

    def _add_subject(self) -> None:
        name = self._clean_text(self.ui.lineEdit_subject_name.text(), default="Subject Name")
        subject_type = self.ui.comboBox_subject_type.currentText()
        format_ = self.ui.comboBox_subject_lec_sem.currentText()
        semester_text = self.ui.comboBox_subject_semester.currentText()
        capacity_text = self._clean_text(self.ui.lineEdit_subject_cap.text(), default="Class Capacity")
        professor_id = self._clean_text(self.ui.lineEdit_subject_prof_id.text(), default="Professors Id").upper()

        if not name:
            self._show_validation_error("Please enter a subject name.")
            return
        if subject_type not in {"Obligatory", "Elective"}:
            self._show_validation_error("Select a subject type.")
            return
        if format_ not in {"Lecture", "Seminar"}:
            self._show_validation_error("Select Lecture or Seminar.")
            return
        if not semester_text.isdigit():
            self._show_validation_error("Select a semester (1-7).")
            return
        semester = int(semester_text)
        try:
            capacity = int(capacity_text)
        except ValueError:
            self._show_validation_error("Capacity must be a number.")
            return
        if capacity <= 0:
            self._show_validation_error("Capacity must be positive.")
            return
        if professor_id and not any(prof.prof_id == professor_id for prof in self.professors):
            self._show_validation_error(f"Unknown professor ID '{professor_id}'.")
            return

        suggestion = self._suggest_subject_code()
        code, ok = QtWidgets.QInputDialog.getText(
            self,
            "Subject Code",
            "Enter unique subject code:",
            text=suggestion,
        )
        if not ok:
            return
        code = self._clean_text(code).upper()
        if not code:
            self._show_validation_error("Subject code cannot be empty.")
            return
        if any(subject.code == code for subject in self.subjects):
            self._show_validation_error(f"Subject code '{code}' already exists.")
            return
        self.subjects.append(
            Subject(
                code=code,
                name=name,
                subject_type=subject_type,
                semester=semester,
                format_=format_,
                capacity=capacity,
                professor_id=professor_id,
            )
        )
        self.subjects.sort(key=lambda subject: subject.code)
        self._refresh_subject_table()
        self._select_table_row(self.ui.tableWidget_subjects, self.subjects, "code", code)
        self._save_subjects()
        self.statusBar().showMessage(f"Subject {code} added.", 4000)

    def _edit_subject(self) -> None:
        row = self._selected_row(self.ui.tableWidget_subjects)
        if row is None:
            self._show_validation_error("Select a subject to edit.")
            return
        subject = self.subjects[row]
        name = self._clean_text(self.ui.lineEdit_subject_name.text()) or subject.name
        subject_type = self.ui.comboBox_subject_type.currentText()
        format_ = self.ui.comboBox_subject_lec_sem.currentText()
        semester_text = self.ui.comboBox_subject_semester.currentText()
        capacity_text = self._clean_text(self.ui.lineEdit_subject_cap.text()) or str(subject.capacity)
        professor_id = self._clean_text(self.ui.lineEdit_subject_prof_id.text()).upper()
        if subject_type not in {"Obligatory", "Elective"}:
            self._show_validation_error("Select a subject type.")
            return
        if format_ not in {"Lecture", "Seminar"}:
            self._show_validation_error("Select Lecture or Seminar.")
            return
        if not semester_text.isdigit():
            self._show_validation_error("Select a semester (1-7).")
            return
        try:
            capacity = int(capacity_text)
        except ValueError:
            self._show_validation_error("Capacity must be a number.")
            return
        if capacity <= 0:
            self._show_validation_error("Capacity must be positive.")
            return
        if professor_id and not any(prof.prof_id == professor_id for prof in self.professors):
            self._show_validation_error(f"Unknown professor ID '{professor_id}'.")
            return
        subject.name = name
        subject.subject_type = subject_type
        subject.format_ = format_
        subject.semester = int(semester_text)
        subject.capacity = capacity
        subject.professor_id = professor_id
        self.subjects.sort(key=lambda subj: subj.code)
        self._refresh_subject_table()
        self._select_table_row(self.ui.tableWidget_subjects, self.subjects, "code", subject.code)
        self._save_subjects()
        self.statusBar().showMessage(f"Subject {subject.code} updated.", 4000)

    def _delete_subject(self) -> None:
        row = self._selected_row(self.ui.tableWidget_subjects)
        if row is None:
            self._show_validation_error("Select a subject to delete.")
            return
        subject = self.subjects[row]
        dependent_schedule = sum(1 for entry in self.schedule_entries if entry.subject_code == subject.code)
        warning = f"Delete subject {subject.code}?\nReferenced by {dependent_schedule} schedule entry(ies)."
        reply = QtWidgets.QMessageBox.question(self, "Confirm Deletion", warning)
        if reply != QtWidgets.QMessageBox.Yes:
            return
        del self.subjects[row]
        self._refresh_subject_table()
        self._save_subjects()
        self.statusBar().showMessage("Subject removed.", 4000)

    def _save_subjects(self) -> None:
        self._write_csv(
            self.subjects_path,
            [subject.as_dict() for subject in self.subjects],
            fieldnames=["SUBJECT_CODE", "SUBJECT_NAME", "TYPE", "SEMESTER", "LEC_SEM", "CLASS_CAP", "PROF_ID"],
        )

    def _import_subjects(self) -> None:
        path = self._open_file_dialog("Import Subjects", self.subjects_path)
        if not path:
            return
        self.subjects_path = path
        self._load_subjects()
        self._refresh_subject_table()
        self.statusBar().showMessage(f"Loaded subjects from {path.name}.", 4000)

    # ----------------------------
    # Schedule
    # ----------------------------
    def _load_schedule(self) -> None:
        self.schedule_entries = []
        if not self.schedule_path.exists():
            return
        try:
            with self.schedule_path.open("r", newline="", encoding="utf-8-sig") as handle:
                reader = csv.DictReader(handle)
                for row in reader:
                    self.schedule_entries.append(
                        ScheduleEntry(
                            schedule_id=row.get("SCHEDULE_ID", "").strip(),
                            subject_code=row.get("SUBJECT_CODE", "").strip(),
                            professor_id=row.get("PROF_ID", "").strip(),
                            room_number=row.get("ROOM_NUMBER", "").strip(),
                            day=row.get("DAY", "").strip(),
                            time=row.get("TIME", "").strip(),
                        )
                    )
        except Exception as exc:  # pragma: no cover
            self._show_io_error(self.schedule_path, exc)
        self.schedule_entries.sort(key=lambda entry: entry.schedule_id)

    def _refresh_schedule_filters(self) -> None:
        with QtCore.QSignalBlocker(self.ui.comboBox_choose_prof):
            current_prof = self.ui.comboBox_choose_prof.currentText()
            self.ui.comboBox_choose_prof.clear()
            self.ui.comboBox_choose_prof.addItem("Choose Professor")
            for prof in self.professors:
                self.ui.comboBox_choose_prof.addItem(prof.prof_id)
            if current_prof in {prof.prof_id for prof in self.professors}:
                index = self.ui.comboBox_choose_prof.findText(current_prof)
                self.ui.comboBox_choose_prof.setCurrentIndex(index)

        with QtCore.QSignalBlocker(self.ui.comboBox_choose_room):
            current_room = self.ui.comboBox_choose_room.currentText()
            self.ui.comboBox_choose_room.clear()
            self.ui.comboBox_choose_room.addItem("Choose Room")
            for room in self.rooms:
                self.ui.comboBox_choose_room.addItem(room.number)
            if current_room in {room.number for room in self.rooms}:
                index = self.ui.comboBox_choose_room.findText(current_room)
                self.ui.comboBox_choose_room.setCurrentIndex(index)

    def _refresh_schedule_table(self) -> None:
        headers = ["Schedule ID", "Subject", "Professor", "Room", "Day", "Time"]
        prof_filter = self.ui.comboBox_choose_prof.currentText()
        room_filter = self.ui.comboBox_choose_room.currentText()
        if prof_filter == "Choose Professor":
            prof_filter = None
        if room_filter == "Choose Room":
            room_filter = None
        rows = []
        for entry in self.schedule_entries:
            if prof_filter and entry.professor_id != prof_filter:
                continue
            if room_filter and entry.room_number != room_filter:
                continue
            rows.append(entry.as_row())
        self._populate_table(self.ui.tableWidget_schedule, rows, headers)

    def _generate_schedule(self) -> None:
        if not self.subjects:
            self._show_validation_error("No subjects available to schedule.")
            return
        if not self.professors:
            self._show_validation_error("No professors available to schedule.")
            return
        if not self.rooms:
            self._show_validation_error("No rooms available to schedule.")
            return

        professors_by_id: Dict[str, Professor] = {prof.prof_id: prof for prof in self.professors}
        subjects_by_code: Dict[str, Subject] = {subject.code: subject for subject in self.subjects}

        eligible_subjects: List[Tuple[Subject, Professor]] = []
        skipped_no_prof: List[str] = []
        for subject in self.subjects:
            professor = professors_by_id.get(subject.professor_id)
            if not subject.professor_id or professor is None:
                skipped_no_prof.append(subject.code)
                continue
            eligible_subjects.append((subject, professor))

        if not eligible_subjects:
            self._show_validation_error("No subjects are linked to a valid professor.")
            return

        rooms_by_type: Dict[str, List[Room]] = {}
        for room in sorted(self.rooms, key=lambda r: (r.lec_sem, r.capacity, r.number)):
            key = room.lec_sem or "Unspecified"
            rooms_by_type.setdefault(key, []).append(room)

        professor_slots: Dict[str, List[Tuple[str, str]]] = {}
        professor_slot_lookup: Dict[str, Set[Tuple[str, str]]] = {}
        for prof in self.professors:
            slots = self._build_professor_slots(prof)
            professor_slots[prof.prof_id] = slots
            professor_slot_lookup[prof.prof_id] = set(slots)

        existing_entries = {entry.subject_code: entry for entry in self.schedule_entries}
        taken_schedule_ids = [entry.schedule_id for entry in self.schedule_entries if entry.schedule_id]
        used_prof_slots: Set[Tuple[str, str, str]] = set()
        used_room_slots: Set[Tuple[str, str, str]] = set()
        new_entries: List[ScheduleEntry] = []
        skipped_no_slot: List[str] = []
        skipped_no_room: List[str] = []

        for subject, professor in eligible_subjects:
            slots_for_prof = professor_slots.get(professor.prof_id, [])
            if not slots_for_prof:
                skipped_no_slot.append(subject.code)
                continue

            entry = existing_entries.get(subject.code)
            if entry is None:
                schedule_id = self._suggest_identifier(taken_schedule_ids, prefix="SCH", digits=3)
                taken_schedule_ids.append(schedule_id)
                entry = ScheduleEntry(schedule_id, subject.code, professor.prof_id, "", "", "")
            else:
                if not entry.schedule_id:
                    schedule_id = self._suggest_identifier(taken_schedule_ids, prefix="SCH", digits=3)
                    taken_schedule_ids.append(schedule_id)
                    entry.schedule_id = schedule_id
                entry.subject_code = subject.code
                entry.professor_id = professor.prof_id

            preferred_slot: Optional[Tuple[str, str]] = None
            if entry.day and entry.time:
                slot_key = (entry.day, entry.time)
                if (
                    (professor.prof_id, entry.day, entry.time) not in used_prof_slots
                    and slot_key in professor_slot_lookup.get(professor.prof_id, set())
                ):
                    preferred_slot = slot_key

            if preferred_slot is None:
                preferred_slot = self._next_available_professor_slot(professor.prof_id, slots_for_prof, used_prof_slots)

            if preferred_slot is None:
                skipped_no_slot.append(subject.code)
                continue

            day, time_value = preferred_slot
            tentative_entry = ScheduleEntry(entry.schedule_id, entry.subject_code, professor.prof_id, entry.room_number, day, time_value)
            room_number = self._assign_room(
                tentative_entry,
                day,
                time_value,
                subjects_by_code,
                rooms_by_type,
                used_room_slots,
            )

            if not room_number:
                skipped_no_room.append(subject.code)
                continue

            entry.professor_id = professor.prof_id
            entry.day = day
            entry.time = time_value
            entry.room_number = room_number
            new_entries.append(entry)
            used_prof_slots.add((professor.prof_id, day, time_value))

        if not new_entries:
            self._show_validation_error("Unable to generate schedule with the current data.")
            return

        self.schedule_entries = sorted(new_entries, key=lambda entry: entry.schedule_id)
        self._save_schedule()
        self._refresh_schedule_filters()
        self._refresh_schedule_table()

        summary_parts = [f"Scheduled {len(self.schedule_entries)} subject(s)"]
        skipped_parts = []
        if skipped_no_prof:
            skipped_parts.append(f"{len(skipped_no_prof)} without professor")
        if skipped_no_slot:
            skipped_parts.append(f"{len(skipped_no_slot)} outside working hours")
        if skipped_no_room:
            skipped_parts.append(f"{len(skipped_no_room)} without available room")
        if skipped_parts:
            summary_parts.append("Skipped: " + ", ".join(skipped_parts))
        self.statusBar().showMessage("; ".join(summary_parts) + ".", 6000)

    def _save_schedule(self) -> None:
        self._write_csv(
            self.schedule_path,
            [entry.as_dict() for entry in self.schedule_entries],
            fieldnames=["SCHEDULE_ID", "SUBJECT_CODE", "PROF_ID", "ROOM_NUMBER", "DAY", "TIME"],
        )

    # ----------------------------
    # Helpers
    # ----------------------------
    def _init_theme_controls(self) -> None:
        self.theme_styles: Dict[str, str] = {
            "Default": "",
            "Midnight": (
                "QMainWindow { background-color: #202124; color: #f1f3f4; }"
                " QWidget { background-color: #272a36; color: #f1f3f4; }"
                " QPushButton { background-color: #394253; color: #f9fbff; border-radius: 6px; padding: 6px 12px; }"
                " QPushButton:hover { background-color: #4a5568; }"
                " QLineEdit, QComboBox { background-color: #2f3442; color: #f9fbff; border: 1px solid #4a5568; border-radius: 4px; padding: 4px 8px; }"
                " QTableWidget { background-color: #1f2230; alternate-background-color: #2b3040; gridline-color: #3c4252; color: #f9fbff; }"
                " QTableWidget::item:selected { background-color: #4c6ef5; color: #ffffff; }"
                " QHeaderView::section { background-color: #2b3040; color: #f9fbff; padding: 6px; border: none; }"
                " QTabBar::tab { background: #2f3442; color: #f9fbff; padding: 10px; border: 1px solid #3c4252; }"
                " QTabBar::tab:selected { background: #3c4252; }"
            ),
            "Ocean": (
                "QMainWindow { background-color: #0f1d2c; color: #f0f4f8; }"
                " QWidget { background-color: #152a3b; color: #f0f4f8; }"
                " QPushButton { background-color: #1f5673; color: #f0f4f8; border-radius: 6px; padding: 6px 12px; }"
                " QPushButton:hover { background-color: #28799d; }"
                " QLineEdit, QComboBox { background-color: #173348; color: #f0f4f8; border: 1px solid #28799d; border-radius: 4px; padding: 4px 8px; }"
                " QTableWidget { background-color: #0f2434; alternate-background-color: #173549; gridline-color: #1f5673; color: #f0f4f8; }"
                " QTableWidget::item:selected { background-color: #2b9ed4; color: #ffffff; }"
                " QHeaderView::section { background-color: #173549; color: #f0f4f8; padding: 6px; border: none; }"
                " QTabBar::tab { background: #173348; color: #f0f4f8; padding: 10px; border: 1px solid #1f5673; }"
                " QTabBar::tab:selected { background: #1f5673; }"
            ),
            "Forest": (
                "QMainWindow { background-color: #14251d; color: #e9f5ec; }"
                " QWidget { background-color: #1c3429; color: #e9f5ec; }"
                " QPushButton { background-color: #22543d; color: #e9f5ec; border-radius: 6px; padding: 6px 12px; }"
                " QPushButton:hover { background-color: #2f6b50; }"
                " QLineEdit, QComboBox { background-color: #1a2f24; color: #e9f5ec; border: 1px solid #2f6b50; border-radius: 4px; padding: 4px 8px; }"
                " QTableWidget { background-color: #112018; alternate-background-color: #1a3527; gridline-color: #2f6b50; color: #e9f5ec; }"
                " QTableWidget::item:selected { background-color: #37966f; color: #ffffff; }"
                " QHeaderView::section { background-color: #1a3527; color: #e9f5ec; padding: 6px; border: none; }"
                " QTabBar::tab { background: #1a2f24; color: #e9f5ec; padding: 10px; border: 1px solid #2f6b50; }"
                " QTabBar::tab:selected { background: #2f6b50; }"
            ),
            "Blossom": (
                "QMainWindow { background-color: #faf5ff; color: #2d1b46; }"
                " QWidget { background-color: #f3e7ff; color: #2d1b46; }"
                " QPushButton { background-color: #c084fc; color: #1b1130; border-radius: 6px; padding: 6px 12px; }"
                " QPushButton:hover { background-color: #a855f7; color: #f9f5ff; }"
                " QLineEdit, QComboBox { background-color: #ffffff; color: #2d1b46; border: 1px solid #c084fc; border-radius: 4px; padding: 4px 8px; }"
                " QTableWidget { background-color: #f1e7ff; alternate-background-color: #e4d4ff; gridline-color: #bb86fc; color: #2d1b46; }"
                " QTableWidget::item:selected { background-color: #b37cfa; color: #ffffff; }"
                " QHeaderView::section { background-color: #c084fc; color: #1b1130; padding: 6px; border: none; }"
                " QTabBar::tab { background: #ede4ff; color: #2d1b46; padding: 10px; border: 1px solid #c084fc; }"
                " QTabBar::tab:selected { background: #b37cfa; color: #f9f5ff; }"
            ),
            "Aurora": (
                "QMainWindow { background-color: #1b1f3b; color: #f0f9ff; }"
                " QWidget { background-color: #232753; color: #f0f9ff; }"
                " QPushButton { background-color: #4c51bf; color: #f0f9ff; border-radius: 6px; padding: 6px 12px; }"
                " QPushButton:hover { background-color: #6366f1; }"
                " QLineEdit, QComboBox { background-color: #2f3469; color: #f0f9ff; border: 1px solid #6366f1; border-radius: 4px; padding: 4px 8px; }"
                " QTableWidget { background-color: #171a36; alternate-background-color: #232b48; gridline-color: #4c51bf; color: #f0f9ff; }"
                " QTableWidget::item:selected { background-color: #7288ff; color: #ffffff; }"
                " QHeaderView::section { background-color: #232b48; color: #f0f9ff; padding: 6px; border: none; }"
                " QTabBar::tab { background: #2f3469; color: #f0f9ff; padding: 10px; border: 1px solid #4c51bf; }"
                " QTabBar::tab:selected { background: #4c51bf; }"
            ),
            "Sunset": (
                "QMainWindow { background-color: #fff7ed; color: #3c1102; }"
                " QWidget { background-color: #ffe7cc; color: #3c1102; }"
                " QPushButton { background-color: #f97316; color: #fff7ed; border-radius: 6px; padding: 6px 12px; }"
                " QPushButton:hover { background-color: #ea580c; color: #fff7ed; }"
                " QLineEdit, QComboBox { background-color: #fff7ed; color: #3c1102; border: 1px solid #fb923c; border-radius: 4px; padding: 4px 8px; }"
                " QTableWidget { background-color: #fff0db; alternate-background-color: #ffdaba; gridline-color: #f97316; color: #3c1102; }"
                " QTableWidget::item:selected { background-color: #f97316; color: #fff7ed; }"
                " QHeaderView::section { background-color: #fb923c; color: #2d0901; padding: 6px; border: none; }"
                " QTabBar::tab { background: #ffedd5; color: #3c1102; padding: 10px; border: 1px solid #fb923c; }"
                " QTabBar::tab:selected { background: #f97316; color: #fff7ed; }"
            ),
            "Slate": (
                "QMainWindow { background-color: #f1f5f9; color: #1f2937; }"
                " QWidget { background-color: #e2e8f0; color: #1f2937; }"
                " QPushButton { background-color: #64748b; color: #f8fafc; border-radius: 6px; padding: 6px 12px; }"
                " QPushButton:hover { background-color: #475569; }"
                " QLineEdit, QComboBox { background-color: #ffffff; color: #1f2937; border: 1px solid #94a3b8; border-radius: 4px; padding: 4px 8px; }"
                " QTableWidget { background-color: #e6edf5; alternate-background-color: #d5deeb; gridline-color: #94a3b8; color: #1f2937; }"
                " QTableWidget::item:selected { background-color: #1d4ed8; color: #ffffff; }"
                " QHeaderView::section { background-color: #c6d2e3; color: #1f2937; padding: 6px; border: none; }"
                " QTabBar::tab { background: #e2e8f0; color: #1f2937; padding: 10px; border: 1px solid #94a3b8; }"
                " QTabBar::tab:selected { background: #1d4ed8; color: #f8fafc; }"
            ),
        }

        self.dark_theme_names: Set[str] = {"Midnight", "Ocean", "Forest", "Aurora"}
        self.default_theme = "Slate"

        self.theme_button = QtWidgets.QToolButton()
        self.theme_button.setText("Theme")
        self.theme_button.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        self.theme_button.setCursor(QtCore.Qt.PointingHandCursor)
        self.theme_button.setToolTip("Change application color theme")
        self.theme_menu = QtWidgets.QMenu(self.theme_button)
        action_group = QtWidgets.QActionGroup(self.theme_menu)
        action_group.setExclusive(True)
        for theme_name in self.theme_styles:
            action = self.theme_menu.addAction(theme_name)
            action.setCheckable(True)
            if theme_name == self.default_theme:
                action.setChecked(True)
            action.triggered.connect(lambda checked, name=theme_name: self._apply_theme(name))
            action_group.addAction(action)
        self.theme_button.setMenu(self.theme_menu)
        self.statusBar().addPermanentWidget(self.theme_button)

        self.current_theme = self.default_theme
        self._apply_theme(self.current_theme, announce=False)

    def _apply_theme(self, name: str, announce: bool = True) -> None:
        app = QtWidgets.QApplication.instance()
        if app is None:
            return
        stylesheet = self.theme_styles.get(name, "")
        app.setStyleSheet(stylesheet)
        self.current_theme = name
        if hasattr(self, "theme_menu"):
            for action in self.theme_menu.actions():
                if action.isCheckable():
                    action.setChecked(action.text() == name)
        self._set_control_palette(name in self.dark_theme_names)
        if announce:
            self.statusBar().showMessage(f"{name} theme applied.", 3000)

    def _set_control_palette(self, use_dark_style: bool) -> None:
        if not hasattr(self, "theme_button"):
            return
        if use_dark_style:
            common_style = (
                "QToolButton { padding: 4px 12px; border-radius: 6px; background-color: #4b5563; color: #f9fafb; }"
                " QToolButton:hover { background-color: #374151; }"
                " QToolButton:pressed { background-color: #1f2937; }"
            )
        else:
            common_style = (
                "QToolButton { padding: 4px 12px; border-radius: 6px; background-color: #e5e7eb; color: #1f2937; border: 1px solid #d1d5db; }"
                " QToolButton:hover { background-color: #d1d5db; }"
                " QToolButton:pressed { background-color: #9ca3af; color: #f9fafb; }"
            )

        self.theme_button.setStyleSheet(common_style)

    def _populate_table(self, table: QtWidgets.QTableWidget, rows: Sequence[Sequence[str]], headers: Sequence[str]) -> None:
        table.setSortingEnabled(False)
        table.clearContents()
        table.setRowCount(len(rows))
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                item = QtWidgets.QTableWidgetItem(value)
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                table.setItem(row_index, column_index, item)
        table.resizeColumnsToContents()
        table.setSortingEnabled(True)

    def _selected_row(self, table: QtWidgets.QTableWidget) -> Optional[int]:
        selection = table.selectionModel()
        if not selection or not selection.hasSelection():
            return None
        return selection.selectedRows()[0].row()

    def _select_table_row(self, table: QtWidgets.QTableWidget, items: Sequence[object], attr: str, value: str) -> None:
        for index, item in enumerate(items):
            if getattr(item, attr) == value:
                table.selectRow(index)
                table.scrollTo(table.model().index(index, 0))
                return
        table.clearSelection()

    def _clean_text(self, text: str, default: str = "") -> str:
        if text is None:
            return ""
        text = text.strip()
        if default and text == default:
            return ""
        return text

    def _open_file_dialog(self, title: str, current: Path) -> Optional[Path]:
        start_dir = str(current.parent if current.exists() else self.data_dir)
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, title, start_dir, "CSV Files (*.csv)")
        if not file_path:
            return None
        return Path(file_path)

    def _write_csv(self, path: Path, rows: Iterable[dict], *, fieldnames: Sequence[str]) -> None:
        try:
            with path.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames)
                writer.writeheader()
                for row in rows:
                    writer.writerow(row)
        except Exception as exc:  # pragma: no cover
            self._show_io_error(path, exc)

    def _show_validation_error(self, message: str) -> None:
        QtWidgets.QMessageBox.warning(self, "Validation", message)

    def _show_missing_file(self, path: Path) -> None:
        QtWidgets.QMessageBox.warning(self, "Missing File", f"Expected data file not found:\n{path}")

    def _show_io_error(self, path: Path, error: Exception) -> None:
        QtWidgets.QMessageBox.critical(self, "File Error", f"Failed to process {path}:\n{error}")

    def _set_combobox_text(self, combo: QtWidgets.QComboBox, text: str) -> None:
        index = combo.findText(text, QtCore.Qt.MatchFixedString)
        if index >= 0:
            combo.setCurrentIndex(index)

    def _suggest_identifier(self, existing: Iterable[str], *, prefix: str, digits: int) -> str:
        pattern = re.compile(rf"{re.escape(prefix)}(\d{{{digits}}})$")
        numbers = [int(match.group(1)) for value in existing if (match := pattern.match(value))]
        next_number = max(numbers, default=0) + 1
        return f"{prefix}{next_number:0{digits}d}"

    def _suggest_subject_code(self) -> str:
        existing = {subject.code for subject in self.subjects}
        prefix = "SBJ"
        counter = 1
        while True:
            candidate = f"{prefix}{counter:03d}"
            if candidate not in existing:
                return candidate
            counter += 1

    def _split_hours(self, text: str) -> Tuple[str, str]:
        if not text:
            return "", ""
        parts = re.findall(r"\d{1,2}(?::\d{2})?", text)
        if len(parts) >= 2:
            start_raw, end_raw = parts[0], parts[1]
        else:
            return "", ""

        def normalize(value: str) -> str:
            if ":" in value:
                hour_text, minute_text = value.split(":", 1)
            else:
                hour_text, minute_text = value, "00"
            try:
                hour = max(0, min(23, int(hour_text)))
                minute = max(0, min(59, int(minute_text)))
            except ValueError:
                return ""
            return f"{hour:02d}:{minute:02d}"

        return normalize(start_raw), normalize(end_raw)

    def _parse_time_string(self, value: str, *, default: time) -> time:
        if not value:
            return default
        candidates = [value.strip(), value.replace(" ", "").strip(), value.replace(".", ":").strip()]
        formats = ("%H:%M", "%H%M", "%H", "%I:%M%p", "%I%p")
        for candidate in filter(None, candidates):
            for fmt in formats:
                try:
                    parsed = datetime.strptime(candidate.upper(), fmt)
                except ValueError:
                    continue
                return time(parsed.hour, parsed.minute)
        return default

    def _normalize_time_range(self, start: time, end: time) -> Tuple[time, time]:
        if start == end:
            base = datetime(1900, 1, 1, start.hour, start.minute)
            end_time = base + timedelta(hours=1)
            return start, time(end_time.hour, end_time.minute)
        if start > end:
            return end, start
        return start, end

    def _time_from_qtime(self, value: QtCore.QTime) -> time:
        return time(value.hour(), value.minute())

    def _qtime_from_time(self, value: time) -> QtCore.QTime:
        return QtCore.QTime(value.hour, value.minute)

    def _build_professor_slots(self, professor: Professor) -> List[Tuple[str, str]]:
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        base_date = datetime.today().date()
        start_dt = datetime.combine(base_date, professor.office_start)
        end_dt = datetime.combine(base_date, professor.office_end)
        if start_dt >= end_dt:
            return []
        slots: List[Tuple[str, str]] = []
        step = timedelta(hours=1)
        for day in days:
            current = start_dt
            while current + step <= end_dt:
                slots.append((day, current.strftime("%H:%M")))
                current += step
        return slots

    def _next_available_professor_slot(
        self,
        professor_id: str,
        slots: Sequence[Tuple[str, str]],
        used_prof_slots: Set[Tuple[str, str, str]],
    ) -> Optional[Tuple[str, str]]:
        for day, time_value in slots:
            if (professor_id, day, time_value) not in used_prof_slots:
                return day, time_value
        return None

    def _build_timeslots(self, count: int) -> List[Tuple[str, str]]:
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        hours = list(range(8, 22))
        slots: List[Tuple[str, str]] = []
        index = 0
        while len(slots) < count:
            day = days[index % len(days)]
            hour = hours[(index // len(days)) % len(hours)]
            slots.append((day, f"{hour:02d}:00"))
            index += 1
        return slots

    def _assign_room(
        self,
        entry: ScheduleEntry,
        day: str,
        time: str,
        subjects_by_code: Dict[str, Subject],
        rooms_by_type: Dict[str, List[Room]],
        used_slots: Set[Tuple[str, str, str]],
    ) -> str:
        subject = subjects_by_code.get(entry.subject_code)
        preferred_type = subject.format_ if subject and subject.format_ else None
        required_capacity = subject.capacity if subject else 0

        current_room = entry.room_number.strip() if entry.room_number else ""
        if current_room and (current_room, day, time) not in used_slots:
            used_slots.add((current_room, day, time))
            return current_room

        candidate_types: List[str] = []
        if preferred_type:
            candidate_types.append(preferred_type)
        for key in rooms_by_type.keys():
            if key not in candidate_types:
                candidate_types.append(key)

        for room_type in candidate_types:
            rooms = rooms_by_type.get(room_type, [])
            if not rooms:
                continue
            suitable = [room for room in rooms if room.capacity >= required_capacity]
            if not suitable:
                suitable = rooms
            for room in suitable:
                slot_key = (room.number, day, time)
                if slot_key in used_slots:
                    continue
                used_slots.add(slot_key)
                return room.number

        for rooms in rooms_by_type.values():
            for room in rooms:
                slot_key = (room.number, day, time)
                if slot_key in used_slots:
                    continue
                used_slots.add(slot_key)
                return room.number

        return current_room

    def _update_professor_completer(self) -> None:
        prof_ids = [prof.prof_id for prof in self.professors]
        completer = QtWidgets.QCompleter(prof_ids)
        completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.ui.lineEdit_subject_prof_id.setCompleter(completer)


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle("University Scheduler")
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
