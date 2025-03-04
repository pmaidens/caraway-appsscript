const HEADER_COLUMNS = 1;
const HEADER_ROWS = 2;
const BLURPLE_CLASS_ROWS = 8;
const BLURPLE_LUNCH_ROWS = 2;
const DEFAULT_CLASS_ROWS = 4;
const DEFAULT_LUNCH_ROWS = 0;
const rooms = ["Blurple", "Yellow", "Green", "Red"];

interface ParsedSlot {
  name: string;
  start: {
      hours: number;
      minutes: number;
  };
  end: {
      hours: number;
      minutes: number;
  };
}

function calculateHours(value = "Liz M", date) { // We need the date value in order to be able to refresh
  const overflowHours = calculateOverflowHours(value);
  const scheduleHours = calculateScheduleHours(value);
  return scheduleHours + overflowHours;
}

function calculateOverflowHours(value: string) {
  const overflowSlots = rooms.map(room => {
    return getOverflowForRoom(room);
  }).flat();
  const hours = calculateTime(parseSlots(overflowSlots, value))
  return hours;
}

function getOverflowForRoom(room: string) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(room);
  if (!sheet) return [] as GoogleAppsScript.Spreadsheet.Range[];
  const values = sheet.getDataRange().getValues();  // Retrieve all values as a 2D array
  const schedulesStart = findCellsWithValue(values, "Schedule");

  return schedulesStart.map((start) => {
    const startRow = getOverflowOffset(room, start);
    const startCol = start.col+HEADER_COLUMNS;
    const colCount = 5;
    return sheet.getRange(startRow, startCol, 10, colCount);
  }).map(scheduleRange => {
    const numRows = scheduleRange.getNumRows();
    const numCols = scheduleRange.getNumColumns();
    const cells: GoogleAppsScript.Spreadsheet.Range[] = [];

    for (let row = 1; row <= numRows; row++) {
      for (let col = 1; col <= numCols; col++) {
        cells.push(scheduleRange.getCell(row, col));
      }
    }
    return cells;
  }).flat();
}

function getOverflowOffset(room: string, start: {row: number, col: number}) {
  switch (room) {
    case "Blurple":
      return start.row + HEADER_ROWS + BLURPLE_CLASS_ROWS + BLURPLE_LUNCH_ROWS + BLURPLE_CLASS_ROWS;
    default:
      return start.row + HEADER_ROWS+DEFAULT_CLASS_ROWS + DEFAULT_LUNCH_ROWS + DEFAULT_CLASS_ROWS;
  }
}

function parseSlots(cells: GoogleAppsScript.Spreadsheet.Range[], value: string) {
  const slots = cells.map(cell => {
        const cellValue = cell.getValue();
        if (!cellValue) return;
        const parsedValue = /^([a-z\s]+)(\s\([a-z]+\))?[\(\s]((\d\d?):?(\d\d)?)-((\d\d?):?(\d\d)?)/gim.exec(cellValue);
        if (parsedValue === null) {
          // cell.setBackground("red");
          return undefined;
        }
        return {
          name: parsedValue[1].trim(),
          start: {
            hours: parseInt(parsedValue[4]),
            minutes: parseInt(parsedValue[5] ?? 0)
          },
          end: {
            hours: parseInt(parsedValue[7]),
            minutes: parseInt(parsedValue[8] ?? 0)
          }
        };
  }).filter(slot => slot && slot.name.toLowerCase() === value.toLowerCase());
  
  return slots;
}

function calculateTime(slots: (ParsedSlot | undefined)[]) {
  return slots.reduce((total, slot) => {
    if (!slot) return total;
    return total + ((convertTo24(slot.end.hours) * 60 + slot.end.minutes) - (convertTo24(slot.start.hours) * 60 + slot.start.minutes))/60
  },0)
}

function convertTo24(hour: number) {
  if (hour > 6 && hour <= 12) {
    return hour;
  }
  return hour + 12;
}

function calculateScheduleHours(value: string) { // We need the date value in order to be able to refresh
  const {classSlotCount, lunchSlotCount} = rooms.reduce(({classSlotCount: count, lunchSlotCount: lunchCount}, room) => {
    const {classSlotCount, lunchSlotCount} = getCountForRoom(room, value);
    return {classSlotCount: count + classSlotCount, lunchSlotCount: lunchCount + lunchSlotCount};
  }, {classSlotCount: 0, lunchSlotCount: 0});
  const hours = calculateHoursFromSlots(classSlotCount, lunchSlotCount);
  return hours;
}

function calculateHoursFromSlots(slotCount: number, lunchCount: number) {
  return slotCount * 3 + lunchCount * 1;
}

function getCountForRoom(room: string, value: string) {
  switch (room) {
    case "Blurple":
      const classSlotCount = getCellCounts(room, 0, BLURPLE_CLASS_ROWS, value) + getCellCounts(room, BLURPLE_CLASS_ROWS+BLURPLE_LUNCH_ROWS, BLURPLE_CLASS_ROWS, value);
      const lunchSlotCount = getCellCounts(room, BLURPLE_CLASS_ROWS, BLURPLE_LUNCH_ROWS, value);
      return {classSlotCount, lunchSlotCount};
    default:
      return {
        classSlotCount: getCellCounts(room, 0, DEFAULT_CLASS_ROWS, value) + getCellCounts(room, DEFAULT_CLASS_ROWS+DEFAULT_LUNCH_ROWS, DEFAULT_CLASS_ROWS, value),
        lunchSlotCount: getCellCounts(room, DEFAULT_CLASS_ROWS, DEFAULT_LUNCH_ROWS, value)
      };
  }
}

function getCellCounts(room: string, rowOffset: number, rowCount: number, value: string) {
  if(!room || rowCount <= 0 || !value) return 0;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(room);
  if (!sheet) return 0;
  const values = sheet.getDataRange().getValues();  // Retrieve all values as a 2D array
  const schedulesStart = findCellsWithValue(values, "Schedule");
  
  const schedules = schedulesStart.map((start) => {
    const startRow = start.row+HEADER_ROWS+rowOffset;
    const startCol = start.col+HEADER_COLUMNS
    const colCount = 5;
    return sheet.getRange(startRow, startCol, rowCount, colCount);
  });
  const cells = filterRangeForMergedCells(schedules);
  const slots = cells.reduce((count, schedule) => {
    return count + schedule.reduce((count, cell) => {
      const cellValue = sheet.getRange(cell).getValues()[0][0]
      if (cellValue === value) {
        return count + 1;
      }
      return count;
    }, 0);
  }, 0)
  return slots;
}

function findCellsWithValue(values: string[][], searchValue: string) {
  const matchingCells: {row: number, col: number}[] = [];          // Array to store matching cell locations

  // Loop through rows and columns to find matches
  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      if (values[row][col] === searchValue) {
        matchingCells.push({
          row: row + 1,      // Add 1 because Apps Script uses 1-based indexing
          col: col + 1       // Add 1 for the column index
        });
      }
    }
  }

  return matchingCells; // Return the list of matching cell coordinates
}

function filterRangeForMergedCells(ranges: GoogleAppsScript.Spreadsheet.Range[]) {
  return ranges.map(range => {
    const startRow = range.getRow();
    const startCol = range.getColumn();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
  
    const validCells: string[] = [];

    const mergedCells = getMergedCellMap(range);
  
    // Iterate through all cells in the range
    for (let row = startRow; row < startRow + numRows; row++) {
      for (let col = startCol; col < startCol + numCols; col++) {
        if (mergedCells[row]?.[col] === undefined) {
          validCells.push(range.getCell(row - startRow + 1, col - startCol + 1).getA1Notation());
        }
      }
    }
  
    return validCells;
  })
}

function getMergedCellMap(range: GoogleAppsScript.Spreadsheet.Range) {
  const mergedRanges = range.getMergedRanges(); // Get all merged ranges in the sheet
  const map: Record<number, Record<number, boolean>> = {};
  for (let i = 0; i < mergedRanges.length; i++) {
    const mergedRange = mergedRanges[i];

    const startRow = mergedRange.getRow();
    const startCol = mergedRange.getColumn();
    const numRows = mergedRange.getNumRows();
    const numCols = mergedRange.getNumColumns();

    for (let row = startRow; row < startRow + numRows; row++) {
      for (let col = startCol; col < startCol + numCols; col++) {
        if (row === startRow && col === startCol) {
          continue;
        }
        map[row] = map[row] ?? {};
        map[row][col] = true;
      }
    }
  }
  return map;
}