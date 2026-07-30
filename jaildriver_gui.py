#!/usr/bin/env python3
"""Desktop search interface for Jaildriver booking-log JSON archives.

Overview
========
The Jaildriver scraper writes one JSON array per day to files named
``booking_logs/YYYY-MM-DD.json``.  This module provides a self-contained PySide6
application that reads those files directly; no database server is required.

The application is split into four main layers:

* :class:`BookingRecord` and :class:`Charge` normalize the JSON data.
* :class:`BookingTableModel` exposes records to Qt's model/view framework.
* :class:`BookingFilterProxyModel` performs compound filtering and sorting.
* :class:`MainWindow` builds the interface, renders details, and exports CSV.

Expected JSON shape
===================
Each daily file should contain a list of objects resembling::

    {
        "booking_id": "...",
        "id": "...",
        "lastname": "...",
        "firstname": "...",
        "middle": "...",
        "sex": "M",
        "dob": "1980-01-31",
        "booked_at": "2026-07-29 13:45:00",
        "charges": [{"code": "...", "name": "..."}]
    }

Missing or malformed scalar fields are normalized to empty strings.  A malformed
file is skipped and reported as a warning so a single damaged archive day does
not prevent the remaining archive from being searched.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable
import subprocess
import os
from PySide6.QtCore import (
    QAbstractTableModel,
    QDate,
    QModelIndex,
    QSettings,
    QSortFilterProxyModel,
    Qt,
    QTimer,
    Signal,
)
from PySide6.QtGui import QAction, QCursor, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTableView,
    QTextBrowser,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

# QSettings uses the organization/application pair as its persistent namespace.
APP_NAME = "Jaildriver Search"
ORGANIZATION = "Frame Analytics Group"
SETTINGS_ARCHIVE_KEY = "archive_directory"


@dataclass(frozen=True)
class Charge:
    """One criminal charge attached to a booking record.

    ``frozen=True`` makes charge objects immutable after parsing.  This prevents
    accidental UI-side mutation of data loaded from the archive.
    """

    code: str
    name: str

    @property
    def display(self) -> str:
        """Return a compact human-readable representation of the charge."""

        code = self.code.strip()
        name = self.name.strip()
        if code and name:
            return f"[{code}] {name}"
        return code or name or "Unknown charge"


@dataclass
class BookingRecord:
    """Normalized booking data used by the models and widgets.

    Dates are parsed up front so filtering and sorting do not repeatedly parse
    strings.  ``repeat_count`` is populated after the complete archive is loaded.
    """

    booking_id: str
    person_id: str
    lastname: str
    firstname: str
    middle: str
    sex: str
    dob: date | None
    booked_at: datetime | None
    charges: list[Charge] = field(default_factory=list)
    source_file: Path | None = None
    repeat_count: int = 1

    @property
    def full_name(self) -> str:
        """Return the person's available name components in display order."""

        parts = [self.firstname, self.middle, self.lastname]
        return " ".join(part.strip() for part in parts if part and part.strip())

    @property
    def age_at_booking(self) -> int | None:
        """Calculate age on the booking date, or ``None`` when unavailable.

        Comparing month/day pairs handles people whose birthday had not yet
        occurred in the booking year.  Negative results indicate inconsistent
        source data and are treated as unknown.
        """

        if self.dob is None or self.booked_at is None:
            return None
        booked = self.booked_at.date()
        years = booked.year - self.dob.year
        if (booked.month, booked.day) < (self.dob.month, self.dob.day):
            years -= 1
        return years if years >= 0 else None

    @property
    def charge_count(self) -> int:
        """Return the number of charge entries on this booking."""

        return len(self.charges)

    @property
    def charge_codes(self) -> str:
        """Return all non-empty charge codes as comma-separated text."""

        return ", ".join(charge.code for charge in self.charges if charge.code)

    @property
    def charge_summary(self) -> str:
        """Return all charges as a semicolon-separated description."""

        return "; ".join(charge.display for charge in self.charges)

    @property
    def identity_key(self) -> tuple[str, ...]:
        """Return the stable key used to identify repeat bookings.

        The source ``person_id`` is preferred because it is less ambiguous than
        a name.  When it is absent, normalized name and date-of-birth fields form
        a conservative fallback identity.
        """

        if self.person_id:
            return ("id", self.person_id.casefold())
        return (
            "person",
            self.lastname.casefold(),
            self.firstname.casefold(),
            self.middle.casefold(),
            self.dob.isoformat() if self.dob else "",
        )

    @property
    def searchable_text(self) -> str:
        """Return a pre-normalized haystack for the global search box.

        ``casefold`` is used instead of ``lower`` for more robust case-insensitive
        comparisons.  The text is generated as a property because records are
        modest in size and it keeps the stored model simple.
        """

        values = [
            self.booking_id,
            self.person_id,
            self.lastname,
            self.firstname,
            self.middle,
            self.full_name,
            self.sex,
            self.dob.isoformat() if self.dob else "",
            self.booked_at.isoformat(sep=" ") if self.booked_at else "",
            self.charge_codes,
            self.charge_summary,
        ]
        return " ".join(values).casefold()


@dataclass
class FilterState:
    """Snapshot of every filter control in :class:`FilterPanel`.

    Keeping filter values in one plain data object decouples the proxy model from
    widgets and makes the acceptance rules straightforward to test.
    """

    global_text: str = ""
    first_name: str = ""
    last_name: str = ""
    booking_id: str = ""
    person_id: str = ""
    sex: str = "All"
    charge_text: str = ""
    use_booking_from: bool = False
    booking_from: date | None = None
    use_booking_to: bool = False
    booking_to: date | None = None
    use_dob_from: bool = False
    dob_from: date | None = None
    use_dob_to: bool = False
    dob_to: date | None = None
    use_min_age: bool = False
    min_age: int = 0
    use_max_age: bool = False
    max_age: int = 120
    use_min_charges: bool = False
    min_charges: int = 0
    use_max_charges: bool = False
    max_charges: int = 99
    repeats_only: bool = False


class ArchiveLoadError(RuntimeError):
    """Raised when the selected archive path itself cannot be loaded."""


def parse_iso_date(value: Any) -> date | None:
    """Parse a date value accepted from historical Jaildriver archives.

    The scraper normally emits ISO dates, but slash- and dash-separated U.S.
    formats are accepted to make the UI tolerant of older exported data.
    """

    text = str(value or "").strip()
    if not text:
        return None
    for pattern in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def parse_booking_datetime(value: Any) -> datetime | None:
    """Parse common booking timestamp formats into :class:`datetime`.

    Explicit formats are tried first for predictable behavior.  ``fromisoformat``
    is the final fallback and also accepts timestamps containing ``T``.
    """

    text = str(value or "").strip()
    if not text:
        return None
    for pattern in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
    ):
        try:
            return datetime.strptime(text, pattern)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def clean_text(value: Any) -> str:
    """Convert a scalar to trimmed text with internal whitespace collapsed."""

    return " ".join(str(value or "").split())


def parse_charges(value: Any) -> list[Charge]:
    """Normalize a JSON charge list into :class:`Charge` objects.

    Dictionary entries are the current format.  Bare strings are accepted for
    compatibility with simpler or older archives.  Other entry types are ignored.
    """

    charges: list[Charge] = []
    if not isinstance(value, list):
        return charges
    for item in value:
        if isinstance(item, dict):
            charges.append(
                Charge(
                    code=clean_text(item.get("code")),
                    name=clean_text(item.get("name")),
                )
            )
        elif isinstance(item, str):
            charges.append(Charge(code="", name=clean_text(item)))
    return charges


def load_archive(directory: Path) -> tuple[list[BookingRecord], list[str]]:
    """Load all YYYY-MM-DD.json files in *directory*.

    Returns ``(records, warnings)``. A malformed file is reported as a warning
    and skipped so one damaged archive day does not block the whole interface.
    """

    if not directory.exists():
        raise ArchiveLoadError(f"Archive directory does not exist: {directory}")
    if not directory.is_dir():
        raise ArchiveLoadError(f"Archive path is not a directory: {directory}")

    # Restrict discovery to the scraper's daily filename convention.  Sorting
    # paths makes warning order deterministic across operating systems.
    paths = sorted(directory.glob("????-??-??.json"))
    records: list[BookingRecord] = []
    warnings: list[str] = []

    for path in paths:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, list):
                raise ValueError("top-level JSON value is not a list")

            for row_number, item in enumerate(data, start=1):
                # A bad row is isolated just like a bad file: retain all usable
                # records rather than treating the archive as transactional.
                if not isinstance(item, dict):
                    warnings.append(f"{path.name}:{row_number}: booking is not an object")
                    continue
                records.append(
                    BookingRecord(
                        booking_id=clean_text(item.get("booking_id")),
                        person_id=clean_text(item.get("id")),
                        lastname=clean_text(item.get("lastname")),
                        firstname=clean_text(item.get("firstname")),
                        middle=clean_text(item.get("middle")),
                        sex=clean_text(item.get("sex")).upper(),
                        dob=parse_iso_date(item.get("dob")),
                        booked_at=parse_booking_datetime(item.get("booked_at")),
                        charges=parse_charges(item.get("charges")),
                        source_file=path,
                    )
                )
        except (OSError, ValueError, json.JSONDecodeError) as exc:
            warnings.append(f"{path.name}: {exc}")

    # Repeat counts are computed only after all daily files are loaded so the
    # metric spans the entire selected archive, not merely one JSON file.
    identity_counts = Counter(record.identity_key for record in records)
    for record in records:
        record.repeat_count = identity_counts[record.identity_key]

    # Present newest bookings first before the user applies any table sorting.
    records.sort(
        key=lambda record: (
            record.booked_at or datetime.min,
            record.lastname.casefold(),
            record.firstname.casefold(),
        ),
        reverse=True,
    )
    return records, warnings


class BookingTableModel(QAbstractTableModel):
    """Read-only table model that exposes :class:`BookingRecord` objects to Qt.

    Qt requests cells by row, column, and role.  DisplayRole returns formatted
    text, UserRole returns naturally sortable values, and RECORD_ROLE exposes the
    underlying record to code that needs more than a single displayed field.
    """

    # Custom role kept separate from Qt.UserRole, which is reserved for sort data.
    RECORD_ROLE = Qt.ItemDataRole.UserRole + 1

    # Each tuple contains the visible heading and BookingRecord attribute name.
    COLUMNS: tuple[tuple[str, str], ...] = (
        ("Booked", "booked_at"),
        ("Name", "full_name"),
        ("Sex", "sex"),
        ("DOB", "dob"),
        ("Age", "age_at_booking"),
        ("Booking ID", "booking_id"),
        ("Person ID", "person_id"),
        ("Bookings", "repeat_count"),
        ("Charges", "charge_count"),
        ("Charge Codes", "charge_codes"),
    )

    def __init__(self, parent: QWidget | None = None) -> None:
        """Initialize an empty model owned by ``parent``."""

        super().__init__(parent)
        self.records: list[BookingRecord] = []

    def set_records(self, records: Iterable[BookingRecord]) -> None:
        """Replace all model rows and notify attached Qt views safely."""

        self.beginResetModel()
        self.records = list(records)
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        """Return top-level row count; this table has no child indexes."""

        return 0 if parent.isValid() else len(self.records)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        """Return the number of declared table columns."""

        return 0 if parent.isValid() else len(self.COLUMNS)

    def headerData(  # noqa: N802
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.ItemDataRole.DisplayRole,
    ) -> Any:
        """Return column labels for the horizontal table header."""

        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal and 0 <= section < len(self.COLUMNS):
            return self.COLUMNS[section][0]
        return section + 1

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        """Return cell content appropriate for the role requested by Qt."""

        if not index.isValid() or not (0 <= index.row() < len(self.records)):
            return None
        record = self.records[index.row()]
        field_name = self.COLUMNS[index.column()][1]
        value = getattr(record, field_name)

        # Qt asks for several roles per cell.  Keeping role-specific values here
        # avoids duplicating display, sorting, tooltip, and selection models.
        if role == self.RECORD_ROLE:
            return record
        if role == Qt.ItemDataRole.TextAlignmentRole:
            if field_name in {"age_at_booking", "repeat_count", "charge_count"}:
                return int(Qt.AlignmentFlag.AlignCenter)
            if field_name == "sex":
                return int(Qt.AlignmentFlag.AlignCenter)
        if role == Qt.ItemDataRole.ToolTipRole:
            return record.charge_summary or "No charges listed"
        if role == Qt.ItemDataRole.UserRole:
            # Return values that sort naturally rather than their display text.
            if isinstance(value, datetime):
                return value.timestamp()
            if isinstance(value, date):
                return value.toordinal()
            return value if value is not None else ""
        if role != Qt.ItemDataRole.DisplayRole:
            return None

        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d %H:%M")
        if isinstance(value, date):
            return value.isoformat()
        if value is None:
            return ""
        return str(value)

    def record_at(self, row: int) -> BookingRecord | None:
        """Return a record by source-model row without raising IndexError."""

        if 0 <= row < len(self.records):
            return self.records[row]
        return None


class BookingFilterProxyModel(QSortFilterProxyModel):
    """Proxy model implementing all interactive search constraints.

    The source model remains unchanged while this model presents only accepted
    rows.  Qt also performs sorting here, allowing filters and sort order to be
    changed without rebuilding the underlying record list.
    """

    filter_state_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = FilterState()
        self.setDynamicSortFilter(True)
        self.setSortRole(Qt.ItemDataRole.UserRole)
        self.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

    def update_state(self, state: FilterState) -> None:
        """Install a new filter snapshot and ask Qt to re-evaluate every row."""

        self.state = state
        self.invalidateFilter()
        self.filter_state_changed.emit()

    @staticmethod
    def _contains(value: str, needle: str) -> bool:
        """Perform an optional case-insensitive substring test."""

        return not needle or needle.casefold() in value.casefold()

    def filterAcceptsRow(  # noqa: N802
        self,
        source_row: int,
        source_parent: QModelIndex,
    ) -> bool:
        """Return ``True`` when one source row satisfies every active filter."""

        model = self.sourceModel()
        if not isinstance(model, BookingTableModel):
            return True
        record = model.record_at(source_row)
        if record is None:
            return False
        state = self.state

        # Space-separated global terms use AND semantics: every term must occur
        # somewhere in the combined searchable record text.
        global_tokens = [token for token in state.global_text.casefold().split() if token]
        if global_tokens and not all(token in record.searchable_text for token in global_tokens):
            return False
        if not self._contains(record.firstname, state.first_name):
            return False
        if not self._contains(record.lastname, state.last_name):
            return False
        if not self._contains(record.booking_id, state.booking_id):
            return False
        if not self._contains(record.person_id, state.person_id):
            return False
        if state.sex != "All" and record.sex != state.sex:
            return False
        if state.charge_text:
            charge_haystack = f"{record.charge_codes} {record.charge_summary}".casefold()
            charge_tokens = [token for token in state.charge_text.casefold().split() if token]
            if not all(token in charge_haystack for token in charge_tokens):
                return False

        # Date bounds are inclusive.  Unknown dates cannot satisfy an active
        # date constraint because their position in the requested range is unknown.
        if state.use_booking_from:
            if record.booked_at is None or record.booked_at.date() < state.booking_from:
                return False
        if state.use_booking_to:
            if record.booked_at is None or record.booked_at.date() > state.booking_to:
                return False
        if state.use_dob_from:
            if record.dob is None or record.dob < state.dob_from:
                return False
        if state.use_dob_to:
            if record.dob is None or record.dob > state.dob_to:
                return False

        age = record.age_at_booking
        if state.use_min_age and (age is None or age < state.min_age):
            return False
        if state.use_max_age and (age is None or age > state.max_age):
            return False
        if state.use_min_charges and record.charge_count < state.min_charges:
            return False
        if state.use_max_charges and record.charge_count > state.max_charges:
            return False
        if state.repeats_only and record.repeat_count < 2:
            return False
        return True


class MetricCard(QFrame):
    """Small reusable dashboard card containing a metric and its label."""

    def __init__(self, title: str, parent: QWidget | None = None) -> None:
        """Create a card with ``title`` and an initial value of zero."""
        super().__init__(parent)
        self.setObjectName("metricCard")
        self.setFrameShape(QFrame.Shape.StyledPanel)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(2)
        self.value_label = QLabel("0")
        self.value_label.setObjectName("metricValue")
        self.title_label = QLabel(title)
        self.title_label.setObjectName("metricTitle")
        layout.addWidget(self.value_label)
        layout.addWidget(self.title_label)

    def set_value(self, value: str | int) -> None:
        """Update the large numeric value displayed by the card."""

        self.value_label.setText(str(value))


class RangeControl(QWidget):
    """Checkbox-controlled date or numeric editor used for optional bounds.

    The checkbox makes a bound active.  Disabling the bound also disables the
    editor but preserves its value for convenient later reactivation.
    """

    changed = Signal()

    def __init__(self, label: str, editor: QWidget, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        self.enabled_box = QCheckBox(label)
        self.editor = editor
        self.editor.setEnabled(False)
        self.enabled_box.toggled.connect(self.editor.setEnabled)
        self.enabled_box.toggled.connect(self.changed)
        # RangeControl accepts either QDateEdit or QSpinBox.  Connect whichever
        # value-change signal the supplied editor exposes.
        if hasattr(self.editor, "dateChanged"):
            self.editor.dateChanged.connect(self.changed)  # type: ignore[attr-defined]
        if hasattr(self.editor, "valueChanged"):
            self.editor.valueChanged.connect(self.changed)  # type: ignore[attr-defined]
        layout.addWidget(self.enabled_box)
        layout.addStretch(1)
        layout.addWidget(self.editor)

    def is_enabled(self) -> bool:
        """Return whether this optional bound should participate in filtering."""

        return self.enabled_box.isChecked()


class FilterPanel(QWidget):
    """Left-hand collection of controls used to construct a FilterState."""

    filters_changed = Signal()
    reset_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._build_ui()

    @staticmethod
    def _date_edit(initial: QDate) -> QDateEdit:
        """Create a consistently configured calendar-backed date editor."""

        editor = QDateEdit(initial)
        editor.setCalendarPopup(True)
        editor.setDisplayFormat("yyyy-MM-dd")
        editor.setMinimumWidth(118)
        return editor

    @staticmethod
    def _spin(minimum: int, maximum: int, value: int) -> QSpinBox:
        """Create a consistently sized integer editor."""

        editor = QSpinBox()
        editor.setRange(minimum, maximum)
        editor.setValue(value)
        editor.setMinimumWidth(84)
        return editor

    def _build_ui(self) -> None:
        """Construct filter groups and wire every control to filters_changed."""

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(12)

        title = QLabel("Search filters")
        title.setObjectName("sectionTitle")
        root.addWidget(title)

        self.global_search = QLineEdit()
        self.global_search.setPlaceholderText("Name, ID, charge, date…")
        self.global_search.setClearButtonEnabled(True)
        root.addWidget(self.global_search)

        # Filters are grouped by domain to keep a large number of fields scannable.
        identity = QGroupBox("Identity")
        identity_form = QFormLayout(identity)
        identity_form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        self.first_name = QLineEdit()
        self.last_name = QLineEdit()
        self.booking_id = QLineEdit()
        self.person_id = QLineEdit()
        self.sex = QComboBox()
        self.sex.addItems(["All", "M", "F", "X", "U"])
        identity_form.addRow("First name", self.first_name)
        identity_form.addRow("Last name", self.last_name)
        identity_form.addRow("Booking ID", self.booking_id)
        identity_form.addRow("Person ID", self.person_id)
        identity_form.addRow("Sex", self.sex)
        root.addWidget(identity)

        charge_group = QGroupBox("Charges")
        charge_layout = QVBoxLayout(charge_group)
        self.charge_text = QLineEdit()
        self.charge_text.setPlaceholderText("Code or description")
        self.charge_text.setClearButtonEnabled(True)
        charge_layout.addWidget(self.charge_text)
        self.min_charges = RangeControl("Minimum", self._spin(0, 99, 1))
        self.max_charges = RangeControl("Maximum", self._spin(0, 99, 10))
        charge_layout.addWidget(self.min_charges)
        charge_layout.addWidget(self.max_charges)
        root.addWidget(charge_group)

        booked_group = QGroupBox("Booking date")
        booked_layout = QVBoxLayout(booked_group)
        today = QDate.currentDate()
        self.booking_from = RangeControl("From", self._date_edit(today.addYears(-1)))
        self.booking_to = RangeControl("Through", self._date_edit(today))
        booked_layout.addWidget(self.booking_from)
        booked_layout.addWidget(self.booking_to)
        root.addWidget(booked_group)

        demographics = QGroupBox("DOB and age at booking")
        demographics_layout = QVBoxLayout(demographics)
        self.dob_from = RangeControl("DOB from", self._date_edit(QDate(1940, 1, 1)))
        self.dob_to = RangeControl("DOB through", self._date_edit(today))
        self.min_age = RangeControl("Minimum age", self._spin(0, 120, 18))
        self.max_age = RangeControl("Maximum age", self._spin(0, 120, 65))
        demographics_layout.addWidget(self.dob_from)
        demographics_layout.addWidget(self.dob_to)
        demographics_layout.addWidget(self.min_age)
        demographics_layout.addWidget(self.max_age)
        root.addWidget(demographics)

        self.repeats_only = QCheckBox("Only people with multiple bookings")
        root.addWidget(self.repeats_only)

        reset_button = QPushButton("Reset filters")
        reset_button.setObjectName("secondaryButton")
        reset_button.clicked.connect(self.reset_requested)
        root.addWidget(reset_button)
        root.addStretch(1)

        line_edits = [
            self.global_search,
            self.first_name,
            self.last_name,
            self.booking_id,
            self.person_id,
            self.charge_text,
        ]
        for widget in line_edits:
            widget.textChanged.connect(self.filters_changed)
        self.sex.currentTextChanged.connect(self.filters_changed)
        self.repeats_only.toggled.connect(self.filters_changed)
        for control in (
            self.min_charges,
            self.max_charges,
            self.booking_from,
            self.booking_to,
            self.dob_from,
            self.dob_to,
            self.min_age,
            self.max_age,
        ):
            control.changed.connect(self.filters_changed)

    @staticmethod
    def _python_date(editor: QDateEdit) -> date:
        """Convert Qt's QDate value into a standard-library date."""

        qdate = editor.date()
        return date(qdate.year(), qdate.month(), qdate.day())

    def state(self) -> FilterState:
        """Capture current widget values in a UI-independent FilterState."""

        return FilterState(
            global_text=self.global_search.text().strip(),
            first_name=self.first_name.text().strip(),
            last_name=self.last_name.text().strip(),
            booking_id=self.booking_id.text().strip(),
            person_id=self.person_id.text().strip(),
            sex=self.sex.currentText(),
            charge_text=self.charge_text.text().strip(),
            use_booking_from=self.booking_from.is_enabled(),
            booking_from=self._python_date(self.booking_from.editor),  # type: ignore[arg-type]
            use_booking_to=self.booking_to.is_enabled(),
            booking_to=self._python_date(self.booking_to.editor),  # type: ignore[arg-type]
            use_dob_from=self.dob_from.is_enabled(),
            dob_from=self._python_date(self.dob_from.editor),  # type: ignore[arg-type]
            use_dob_to=self.dob_to.is_enabled(),
            dob_to=self._python_date(self.dob_to.editor),  # type: ignore[arg-type]
            use_min_age=self.min_age.is_enabled(),
            min_age=self.min_age.editor.value(),  # type: ignore[attr-defined]
            use_max_age=self.max_age.is_enabled(),
            max_age=self.max_age.editor.value(),  # type: ignore[attr-defined]
            use_min_charges=self.min_charges.is_enabled(),
            min_charges=self.min_charges.editor.value(),  # type: ignore[attr-defined]
            use_max_charges=self.max_charges.is_enabled(),
            max_charges=self.max_charges.editor.value(),  # type: ignore[attr-defined]
            repeats_only=self.repeats_only.isChecked(),
        )

    def reset(self) -> None:
        """Clear text filters and deactivate every optional range bound."""

        for widget in (
            self.global_search,
            self.first_name,
            self.last_name,
            self.booking_id,
            self.person_id,
            self.charge_text,
        ):
            widget.clear()
        self.sex.setCurrentIndex(0)
        self.repeats_only.setChecked(False)
        for control in (
            self.min_charges,
            self.max_charges,
            self.booking_from,
            self.booking_to,
            self.dob_from,
            self.dob_to,
            self.min_age,
            self.max_age,
        ):
            control.enabled_box.setChecked(False)
        self.filters_changed.emit()


class MainWindow(QMainWindow):
    """Top-level Jaildriver search window and application controller."""

    def __init__(self, initial_archive: Path | None = None) -> None:
        """Build the window and load the best available initial archive.

        Archive selection precedence is: command-line path, saved QSettings path,
        then a local ``booking_logs`` directory.
        """
        super().__init__()
        # QSettings remembers the last archive between application launches.
        self.settings = QSettings(ORGANIZATION, APP_NAME)
        self.archive_directory: Path | None = None
        self.source_model = BookingTableModel(self)
        self.proxy_model = BookingFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.source_model)
        # TextChanged can fire on every keystroke.  A short single-shot timer
        # batches those events so large archives are not filtered excessively.
        self.filter_timer = QTimer(self)
        self.filter_timer.setSingleShot(True)
        self.filter_timer.setInterval(160)
        self.filter_timer.timeout.connect(self.apply_filters)
        self._build_ui()
        self._build_actions()
        self._apply_stylesheet()

        # Resolve the startup archive without interrupting the user with a file
        # dialog.  They can always choose a different directory from the toolbar.
        candidate = initial_archive
        if candidate is None:
            stored = self.settings.value(SETTINGS_ARCHIVE_KEY, "", str)
            if stored:
                candidate = Path(stored)
        if candidate is None and Path("booking_logs").is_dir():
            candidate = Path("booking_logs")
        if candidate is not None:
            self.load_directory(candidate, show_errors=False)
        else:
            self._show_empty_state()

    def _build_ui(self) -> None:
        """Create the header, metric cards, filter panel, table, and detail pane."""

        self.setWindowTitle(APP_NAME)
        self.resize(1580, 920)
        self.setMinimumSize(1100, 680)

        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(18, 14, 18, 16)
        root.setSpacing(12)
        self.setCentralWidget(central)

        heading_layout = QHBoxLayout()
        heading_text = QVBoxLayout()
        title = QLabel("James' Creepy Private Repository of SLO County Sheriff's Jail Bookings  ")
        title.setObjectName("pageTitle")
        self.subtitle = QLabel("Search normalized SLO Sheriff booking archives")
        self.subtitle.setObjectName("pageSubtitle")
        heading_text.addWidget(title)
        heading_text.addWidget(self.subtitle)
        heading_layout.addLayout(heading_text)
        heading_layout.addStretch(1)
        self.open_button = QPushButton("Open archive")
        self.open_button.clicked.connect(self.choose_directory)
        self.reload_button = QPushButton("Reload")
        self.reload_button.setObjectName("secondaryButton")
        self.reload_button.clicked.connect(self.reload_archive)
        self.export_button = QPushButton("Export filtered CSV")
        self.export_button.clicked.connect(self.export_csv)
        heading_layout.addWidget(self.open_button)
        heading_layout.addWidget(self.reload_button)
        heading_layout.addWidget(self.export_button)
        root.addLayout(heading_layout)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(10)
        self.total_card = MetricCard("Filtered bookings")
        self.people_card = MetricCard("Unique people")
        self.charges_card = MetricCard("Charge entries")
        self.repeats_card = MetricCard("Repeat-booking records")
        metrics.addWidget(self.total_card, 0, 0)
        metrics.addWidget(self.people_card, 0, 1)
        metrics.addWidget(self.charges_card, 0, 2)
        metrics.addWidget(self.repeats_card, 0, 3)
        root.addLayout(metrics)

        # The three panes remain independently resizable: filters, results, details.
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

        self.filter_panel = FilterPanel()
        self.filter_panel.setMinimumWidth(285)
        self.filter_panel.setMaximumWidth(390)
        self.filter_panel.filters_changed.connect(self.schedule_filter)
        self.filter_panel.reset_requested.connect(self.filter_panel.reset)
        filter_scroll = QScrollArea()
        filter_scroll.setWidgetResizable(True)
        filter_scroll.setFrameShape(QFrame.Shape.NoFrame)
        filter_scroll.setWidget(self.filter_panel)
        splitter.addWidget(filter_scroll)

        table_container = QWidget()
        table_layout = QVBoxLayout(table_container)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_layout.setSpacing(7)
        result_heading = QHBoxLayout()
        result_title = QLabel("Results")
        result_title.setObjectName("sectionTitle")
        self.result_count = QLabel("0 records")
        self.result_count.setObjectName("mutedLabel")
        result_heading.addWidget(result_title)
        result_heading.addStretch(1)
        result_heading.addWidget(self.result_count)
        table_layout.addLayout(result_heading)

        self.table = QTableView()
        self.table.setModel(self.proxy_model)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.table.setSortingEnabled(True)
        self.table.setWordWrap(False)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(34)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setMinimumSectionSize(65)
        for column, width in enumerate((145, 205, 48, 95, 55, 130, 105, 75, 68, 220)):
            self.table.setColumnWidth(column, width)
        self.table.selectionModel().selectionChanged.connect(self.show_selected_record)
        self.table.doubleClicked.connect(self.show_selected_record)
        table_layout.addWidget(self.table, 1)
        splitter.addWidget(table_container)

        detail_container = QWidget()
        detail_layout = QVBoxLayout(detail_container)
        detail_layout.setContentsMargins(12, 0, 0, 0)
        detail_layout.setSpacing(7)
        detail_title = QLabel("Booking details")
        detail_title.setObjectName("sectionTitle")
        self.details = QTextBrowser()
        self.details.setOpenExternalLinks(False)
        self.details.setMinimumWidth(300)
        detail_layout.addWidget(detail_title)
        detail_layout.addWidget(self.details, 1)
        splitter.addWidget(detail_container)
        splitter.setSizes([310, 900, 350])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 0)

        status = QStatusBar()
        self.setStatusBar(status)
        self.status_label = QLabel("No archive loaded")
        status.addWidget(self.status_label, 1)

    def _build_actions(self) -> None:
        """Create toolbar commands and standard keyboard shortcuts."""

        toolbar = QToolBar("Archive")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        open_action = QAction("Open archive", self)
        open_action.setShortcut(QKeySequence.StandardKey.Open)
        open_action.triggered.connect(self.choose_directory)
        toolbar.addAction(open_action)

        reload_action = QAction("Reload", self)
        reload_action.setShortcut(QKeySequence.StandardKey.Refresh)
        reload_action.triggered.connect(self.reload_archive)
        toolbar.addAction(reload_action)

        export_action = QAction("Export CSV", self)
        export_action.setShortcut(QKeySequence("Ctrl+Shift+S"))
        export_action.triggered.connect(self.export_csv)
        toolbar.addAction(export_action)

        toolbar.addSeparator()
        focus_search = QAction("Focus search", self)
        focus_search.setShortcut(QKeySequence.StandardKey.Find)
        focus_search.triggered.connect(self.filter_panel.global_search.setFocus)
        toolbar.addAction(focus_search)

        quit_action = QAction("Quit", self)
        quit_action.setShortcut(QKeySequence.StandardKey.Quit)
        quit_action.triggered.connect(self.close)
        self.addAction(quit_action)

    def _apply_stylesheet(self) -> None:
        """Apply the application-wide Qt stylesheet.

        Object names assigned during UI construction allow selected widgets to
        receive specialized styling without custom widget subclasses.
        """

        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background: #f4f6f8;
                color: #17212b;
                font-size: 13px;
            }
            QToolBar {
                background: #ffffff;
                border: none;
                border-bottom: 1px solid #d9e0e7;
                spacing: 6px;
                padding: 5px 10px;
            }
            QStatusBar {
                background: #ffffff;
                border-top: 1px solid #d9e0e7;
            }
            QLabel#pageTitle {
                font-size: 25px;
                font-weight: 700;
                color: #102a43;
            }
            QLabel#pageSubtitle, QLabel#mutedLabel, QLabel#metricTitle {
                color: #627d98;
            }
            QLabel#sectionTitle {
                font-size: 15px;
                font-weight: 700;
                color: #243b53;
            }
            QLabel#metricValue {
                font-size: 24px;
                font-weight: 700;
                color: #0b7285;
            }
            QFrame#metricCard {
                background: #ffffff;
                border: 1px solid #d9e2ec;
                border-radius: 8px;
            }
            QGroupBox {
                background: #ffffff;
                border: 1px solid #d9e2ec;
                border-radius: 7px;
                margin-top: 11px;
                padding-top: 8px;
                font-weight: 600;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #334e68;
            }
            QLineEdit, QComboBox, QDateEdit, QSpinBox {
                background: #ffffff;
                border: 1px solid #bcccdc;
                border-radius: 5px;
                padding: 6px 8px;
                min-height: 20px;
                selection-background-color: #0b7285;
            }
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QSpinBox:focus {
                border: 1px solid #0b7285;
            }
            QPushButton {
                background: #0b7285;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 8px 13px;
                font-weight: 600;
            }
            QPushButton:hover { background: #0a6373; }
            QPushButton:pressed { background: #075b68; }
            QPushButton#secondaryButton {
                background: #e8eef3;
                color: #243b53;
                border: 1px solid #cbd5df;
            }
            QPushButton#secondaryButton:hover { background: #dce5ec; }
            QTableView {
                background: #ffffff;
                alternate-background-color: #f7fafc;
                border: 1px solid #d9e2ec;
                border-radius: 7px;
                selection-background-color: #cce9ed;
                selection-color: #102a43;
            }
            QHeaderView::section {
                background: #eaf0f4;
                color: #243b53;
                border: none;
                border-right: 1px solid #d4dde5;
                border-bottom: 1px solid #cbd5df;
                padding: 8px;
                font-weight: 700;
            }
            QTextBrowser {
                background: #ffffff;
                border: 1px solid #d9e2ec;
                border-radius: 7px;
                padding: 10px;
            }
            QScrollArea { background: transparent; }
            QSplitter::handle { background: #d9e2ec; width: 1px; }
            """
        )

    def _show_empty_state(self) -> None:
        """Reset the interface and explain how to choose an archive."""

        self.source_model.set_records([])
        self.subtitle.setText("Open a booking_logs directory to begin")
        self.details.setHtml(
            "<h3>No archive loaded</h3>"
            "<p>Click <b>Open archive</b> and select the directory containing "
            "<code>YYYY-MM-DD.json</code> files produced by jaildriver.py.</p>"
        )
        self.update_metrics()

    def choose_directory(self) -> None:
        """Prompt for an archive directory and load the selected path."""

        start = str(self.archive_directory or Path.cwd())
        selected = QFileDialog.getExistingDirectory(self, "Open Jaildriver archive", start)
        if selected:
            self.load_directory(Path(selected), show_errors=True)

    def load_directory(self, directory: Path, *, show_errors: bool = True) -> None:
        """Load ``directory`` into the model and refresh every dependent view.

        ``show_errors`` is disabled during automatic startup discovery so an old
        saved path does not immediately display a modal error dialog.
        """

        QApplication.setOverrideCursor(QCursor(Qt.CursorShape.WaitCursor))
        try:
            records, warnings = load_archive(directory)
        except ArchiveLoadError as exc:
            if show_errors:
                QMessageBox.critical(self, "Could not load archive", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        self.archive_directory = directory.resolve()
        self.settings.setValue(SETTINGS_ARCHIVE_KEY, str(self.archive_directory))
        # Updating the source model automatically informs the table; the proxy is
        # explicitly invalidated because its acceptance result depends on each row.
        self.source_model.set_records(records)
        self.proxy_model.invalidateFilter()
        self.proxy_model.sort(0, Qt.SortOrder.DescendingOrder)
        self.subtitle.setText(str(self.archive_directory))
        self.status_label.setText(
            f"Loaded {len(records):,} bookings from {len(list(directory.glob('????-??-??.json'))):,} files"
        )
        self.apply_filters()
        if self.proxy_model.rowCount() > 0:
            self.table.selectRow(0)
        else:
            self.details.setHtml("<h3>No booking records found</h3>")

        if warnings and show_errors:
            preview = "\n".join(warnings[:15])
            suffix = "" if len(warnings) <= 15 else f"\n…and {len(warnings) - 15} more"
            QMessageBox.warning(
                self,
                "Archive loaded with warnings",
                f"The archive loaded, but {len(warnings)} issue(s) were skipped:\n\n{preview}{suffix}",
            )

    def reload_archive(self) -> None:
        """Reload the current path so newly scraped daily files become visible."""

        if self.archive_directory is None:
            self.choose_directory()
            return
        self.load_directory(self.archive_directory, show_errors=True)

    def schedule_filter(self) -> None:
        """Debounce rapid control changes before applying expensive filters."""

        self.filter_timer.start()

    def apply_filters(self) -> None:
        """Apply the panel state, refresh metrics, and select a visible record."""

        self.proxy_model.update_state(self.filter_panel.state())
        self.update_metrics()
        if self.proxy_model.rowCount() > 0:
            self.table.selectRow(0)
        else:
            self.details.setHtml(
                "<h3>No matching records</h3><p>Adjust or reset the filters.</p>"
            )

    def iter_filtered_records(self) -> Iterable[BookingRecord]:
        """Yield records in the proxy model's current filtered and sorted order."""

        for proxy_row in range(self.proxy_model.rowCount()):
            # Proxy rows reflect both filtering and user-selected sort order.
            proxy_index = self.proxy_model.index(proxy_row, 0)
            source_index = self.proxy_model.mapToSource(proxy_index)
            record = self.source_model.record_at(source_index.row())
            if record is not None:
                yield record

    def update_metrics(self) -> None:
        """Recalculate summary cards from the currently visible record set."""

        # Materialize once because all four metrics traverse the same visible set.
        records = list(self.iter_filtered_records())
        unique_people = len({record.identity_key for record in records})
        charge_entries = sum(record.charge_count for record in records)
        repeated_records = sum(1 for record in records if record.repeat_count > 1)
        self.total_card.set_value(f"{len(records):,}")
        self.people_card.set_value(f"{unique_people:,}")
        self.charges_card.set_value(f"{charge_entries:,}")
        self.repeats_card.set_value(f"{repeated_records:,}")
        noun = "record" if len(records) == 1 else "records"
        self.result_count.setText(f"{len(records):,} {noun}")

    def current_record(self) -> BookingRecord | None:
        """Map the selected proxy row back to its source BookingRecord."""

        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return None
        # View indexes belong to the proxy; model data is stored in source rows.
        source_index = self.proxy_model.mapToSource(indexes[0])
        return self.source_model.record_at(source_index.row())

    def show_selected_record(self, *_args: Any) -> None:
        """Render the current table selection as escaped HTML in the detail pane."""

        record = self.current_record()
        if record is None:
            return
        charges_html = "".join(
            f"<tr><td><code>{self._html(charge.code)}</code></td>"
            f"<td>{self._html(charge.name)}</td></tr>"
            for charge in record.charges
        )
        if not charges_html:
            charges_html = '<tr><td colspan="2"><i>No charges listed</i></td></tr>'
        age = "" if record.age_at_booking is None else str(record.age_at_booking)
        source = record.source_file.name if record.source_file else ""
        self.details.setHtml(
            f"""
            <style>
              body {{ color: #243b53; font-family: sans-serif; }}
              h2 {{ margin-bottom: 2px; color: #102a43; }}
              .muted {{ color: #627d98; margin-top: 0; }}
              table.meta {{ border-collapse: collapse; width: 100%; margin: 12px 0; }}
              table.meta td {{ padding: 5px 4px; border-bottom: 1px solid #e3e9ef; }}
              table.meta td:first-child {{ color: #627d98; width: 38%; }}
              table.charges {{ border-collapse: collapse; width: 100%; }}
              table.charges th, table.charges td {{ padding: 7px; border: 1px solid #d9e2ec; text-align: left; }}
              table.charges th {{ background: #eaf0f4; }}
              code {{ color: #0b7285; }}
            </style>
            <h2>{self._html(record.full_name or 'Unnamed record')}</h2>
            <p class="muted">{self._html(record.booking_id)}</p>
            <table class="meta">
              <tr><td>Booked</td><td>{self._html(record.booked_at.strftime('%Y-%m-%d %H:%M:%S') if record.booked_at else '')}</td></tr>
              <tr><td>Person ID</td><td>{self._html(record.person_id)}</td></tr>
              <tr><td>Sex</td><td>{self._html(record.sex)}</td></tr>
              <tr><td>Date of birth</td><td>{self._html(record.dob.isoformat() if record.dob else '')}</td></tr>
              <tr><td>Age at booking</td><td>{age}</td></tr>
              <tr><td>Bookings in archive</td><td>{record.repeat_count}</td></tr>
              <tr><td>Source file</td><td>{self._html(source)}</td></tr>
            </table>
            <h3>Charges ({record.charge_count})</h3>
            <table class="charges">
              <tr><th>Code</th><th>Description</th></tr>
              {charges_html}
            </table>
            """
        )

    @staticmethod
    def _html(value: Any) -> str:
        """Escape archive text before inserting it into QTextBrowser HTML."""

        text = str(value or "")
        # Ampersand must be escaped first so entities introduced by later
        # replacements are not escaped a second time.
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&#39;")
        )

    def export_csv(self) -> None:
        """Write the visible record set to a user-selected UTF-8 CSV file."""

        records = list(self.iter_filtered_records())
        if not records:
            QMessageBox.information(self, "Nothing to export", "No filtered records are available.")
            return
        default_name = "jaildriver-filtered.csv"
        if self.archive_directory:
            default_path = self.archive_directory / default_name
        else:
            default_path = Path.cwd() / default_name
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "Export filtered bookings",
            str(default_path),
            "CSV files (*.csv);;All files (*)",
        )
        if not selected:
            return
        path = Path(selected)
        try:
            with path.open("w", encoding="utf-8", newline="") as output:
                writer = csv.writer(output)
                # Keep the header and row field order synchronized; downstream
                # spreadsheet/database imports rely on these stable column names.
                writer.writerow(
                    [
                        "booked_at",
                        "booking_id",
                        "person_id",
                        "lastname",
                        "firstname",
                        "middle",
                        "full_name",
                        "sex",
                        "dob",
                        "age_at_booking",
                        "bookings_in_archive",
                        "charge_count",
                        "charge_codes",
                        "charges",
                        "source_file",
                    ]
                )
                for record in records:
                    writer.writerow(
                        [
                            record.booked_at.isoformat(sep=" ") if record.booked_at else "",
                            record.booking_id,
                            record.person_id,
                            record.lastname,
                            record.firstname,
                            record.middle,
                            record.full_name,
                            record.sex,
                            record.dob.isoformat() if record.dob else "",
                            record.age_at_booking if record.age_at_booking is not None else "",
                            record.repeat_count,
                            record.charge_count,
                            record.charge_codes,
                            record.charge_summary,
                            record.source_file.name if record.source_file else "",
                        ]
                    )
        except OSError as exc:
            QMessageBox.critical(self, "Export failed", str(exc))
            return
        self.status_label.setText(f"Exported {len(records):,} records to {path}")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse the optional archive directory supplied on the command line."""

    parser = argparse.ArgumentParser(description="Search Jaildriver JSON booking archives with Qt.")
    parser.add_argument(
        "archive_dir",
        nargs="?",
        type=Path,
        help="Directory containing booking_logs/YYYY-MM-DD.json files",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Create QApplication, show the main window, and run Qt's event loop."""

    todays_log = "booking_logs/" + datetime.now().strftime("%Y-%m-%d") + ".json"
    if os.path.islink(__file__):
        here = os.path.dirname(os.readlink(__file__))
    else:
        here = os.path.dirname(__file__)

    today = os.path.join(here, todays_log)

    if not os.path.exists(today):
        yn = input(f"Booking log for {today} does not exist (stale data). Re-crawl the booking website? ")
        if yn.lower() in ["yes", "y"]:
            py = os.path.join(here, ".venv/bin/python3")
            subprocess.run(f"{py} jaildriver.py", shell = True)
    args = parse_args(argv)
    app = QApplication(sys.argv if argv is None else [sys.argv[0], *argv])
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(ORGANIZATION)
    window = MainWindow(args.archive_dir)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
