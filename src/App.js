import React, { useState, useEffect, useRef } from "react";
import "./App.css";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { evaluate } from "mathjs"; // Для вычислений формул

const App = () => {
  const [rows, setRows] = useState(
    Array(20)
      .fill(0)
      .map(() =>
        Array(10).fill({
          value: "",
          style: {},
        })
      )
  );
  const [selectedCell, setSelectedCell] = useState(null); // Хранение выделенной ячейки
  const [formula, setFormula] = useState(""); // Формула для ввода
  const [selectedRange, setSelectedRange] = useState([]); // Для хранения диапазона
  const [isSelectingRange, setIsSelectingRange] = useState(false); // Для контроля за выделением диапазона
  const [ctrlPressed, setCtrlPressed] = useState(false); // Статус клавиши Ctrl
  const tableRef = useRef(null); // Ссылка на таблицу для контроля за координатами

  useEffect(() => {
    const handleKeyDown = (event) => {
      if (event.key === "Control" || event.key === "Meta") {
        setCtrlPressed(true);
      }
      if (event.key === "Delete") {
        const newRows = [...rows];
        if (selectedRange.length === 2) {
          const [startRow, startCol] = selectedRange[0];
          const [endRow, endCol] = selectedRange[1];
          for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
            for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
              newRows[r][c].value = "";
            }
          }
        } else if (selectedCell) {
          newRows[selectedCell.row][selectedCell.col].value = "";
        }
        setRows(newRows);
      }
    };

    const handleKeyUp = (event) => {
      if (event.key === "Control" || event.key === "Meta") {
        setCtrlPressed(false);
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    window.addEventListener("keyup", handleKeyUp);

    return () => {
      window.removeEventListener("keydown", handleKeyDown);
      window.removeEventListener("keyup", handleKeyUp);
    };
  }, [selectedCell, selectedRange, rows]);

  // Преобразует индексы строки и столбца в ссылку на ячейку (например, A1, B2)
  const indexToCell = (row, col) => {
    const colName = String.fromCharCode(65 + col); // Преобразуем индекс колонки в букву (0 -> 'A', 1 -> 'B', ...)
    return `${colName}${row + 1}`;
  };

  // Обновление значения ячейки
  const handleCellChange = (rowIndex, colIndex, value) => {
    const newRows = [...rows];
    newRows[rowIndex][colIndex] = {
      ...newRows[rowIndex][colIndex],
      value,
    };
    setRows(newRows);
  };

  // Обработка клика по ячейке
  const expandGridIfNeeded = (rowIndex, colIndex) => {
    const newRows = [...rows];
    const numRows = newRows.length;
    const numCols = newRows[0].length;

    if (rowIndex >= numRows - 1) {
      // Add a new row
      newRows.push(Array(numCols).fill({ value: "", style: {} }));
    }

    if (colIndex >= numCols - 1) {
      // Add a new column
      newRows.forEach(row => row.push({ value: "", style: {} }));
    }

    setRows(newRows);
  };

  const handleCellClick = (rowIndex, colIndex) => {
    setSelectedCell({ row: rowIndex, col: colIndex });
    setFormula(rows[rowIndex][colIndex].value || "");
    expandGridIfNeeded(rowIndex, colIndex);
  };

  // Начало выделения диапазона
  const startRangeSelection = (rowIndex, colIndex) => {
    setIsSelectingRange(true);
    setSelectedRange([[rowIndex, colIndex], [rowIndex, colIndex]]);
  };

  // Обработка мыши для выделения диапазона
  const handleMouseMove = (e, rowIndex, colIndex) => {
    if (isSelectingRange) {
      const [start, end] = selectedRange;
      const newStart = [
        Math.min(start[0], rowIndex),
        Math.min(start[1], colIndex)
      ];
      const newEnd = [
        Math.max(start[0], rowIndex),
        Math.max(start[1], colIndex)
      ];
      setSelectedRange([newStart, newEnd]);
    }
  };

  // Завершение выделения диапазона
  const handleMouseUp = () => {
    setIsSelectingRange(false); // Завершаем выделение
  };

  // Обработка изменения формулы
  const handleFormulaChange = (e) => {
    setFormula(e.target.value);
    if (selectedCell && selectedRange.length === 0) {
      const newRows = [...rows];
      newRows[selectedCell.row][selectedCell.col].value = e.target.value;
      setRows(newRows);
    }
  };

  // Применение формулы при нажатии Enter
  const handleKeyDown = (e) => {
    if (e.key === 'Enter' && selectedCell) {
      e.preventDefault();
      const { row, col } = selectedCell;
      const currentValue = formula;
      
      handleCellChange(row, col, currentValue);
      
      if (currentValue.startsWith('=')) {
        applyFormula(row, col);
      }
    }
  };

  // Применение формулы
  const applyFormula = (rowIndex, colIndex) => {
    const value = rows[rowIndex]?.[colIndex]?.value;

    if (typeof value === 'string' && value.startsWith('=')) {
      try {
        // Убираем "=" и рассчитываем формулу
        let finalFormula = value.slice(1);

        // Обработка функций
        finalFormula = handleFunctions(finalFormula);

        // Заменяем ссылки на значения
        finalFormula = finalFormula.replace(/[A-Z][0-9]+/g, (cellRef) => {
          const cellValue = getCellValue(cellRef);
          return cellValue !== undefined ? cellValue : '0';
        });

        // Вычисляем результат и обновляем ячейку
        const result = evaluate(finalFormula);
        handleCellChange(rowIndex, colIndex, result.toString());
        setFormula(result.toString());
      } catch (error) {
        console.error("Ошибка формулы:", error);
        handleCellChange(rowIndex, colIndex, "ERROR");
        setFormula("ERROR");
      }
    }
  };

  // Обработка функций в формулах
  const handleFunctions = (formula) => {
    const functions = {
      SUM: (range) => {
        const [start, end] = range.split(':');
        const startCell = parseCellReference(start);
        const endCell = parseCellReference(end);
        let sum = 0;

        for (let r = startCell.row; r <= endCell.row; r++) {
          for (let c = startCell.col; c <= endCell.col; c++) {
            sum += parseFloat(rows[r][c].value) || 0;
          }
        }
        return sum;
      },
      MIN: (range) => {
        const [start, end] = range.split(':');
        const startCell = parseCellReference(start);
        const endCell = parseCellReference(end);
        let min = Infinity;

        for (let r = startCell.row; r <= endCell.row; r++) {
          for (let c = startCell.col; c <= endCell.col; c++) {
            const value = parseFloat(rows[r][c].value);
            if (!isNaN(value) && value < min) {
              min = value;
            }
          }
        }
        return min;
      },
      MAX: (range) => {
        const [start, end] = range.split(':');
        const startCell = parseCellReference(start);
        const endCell = parseCellReference(end);
        let max = -Infinity;

        for (let r = startCell.row; r <= endCell.row; r++) {
          for (let c = startCell.col; c <= endCell.col; c++) {
            const value = parseFloat(rows[r][c].value);
            if (!isNaN(value) && value > max) {
              max = value;
            }
          }
        }
        return max;
      },
      AVRG: (range) => {
        const [start, end] = range.split(':');
        const startCell = parseCellReference(start);
        const endCell = parseCellReference(end);
        let sum = 0;
        let count = 0;

        for (let r = startCell.row; r <= endCell.row; r++) {
          for (let c = startCell.col; c <= endCell.col; c++) {
            sum += parseFloat(rows[r][c].value) || 0;
            count++;
          }
        }
        return count ? sum / count : 0;
      },
    };

    return formula.replace(/(SUM|MIN|MAX|AVRG)\(([^)]+)\)/g, (match, func, range) => {
      const fn = functions[func];
      return fn ? fn(range) : match;
    });
  };

  // Преобразование ссылки на ячейку в индексы
  const parseCellReference = (cellRef) => {
    const col = cellRef.charCodeAt(0) - 65;
    const row = parseInt(cellRef.slice(1)) - 1;
    return { row, col };
  };

  // Преобразование ссылки на ячейку в значение
  const getCellValue = (cellRef) => {
    const col = cellRef.charCodeAt(0) - 65;
    const row = parseInt(cellRef.slice(1)) - 1;
    return parseFloat(rows[row]?.[col]?.value) || 0;
  };

  // Выбор диапазона без потери фокуса
  const handleMouseEnter = (rowIndex, colIndex) => {
    if (isSelectingRange) {
      setSelectedRange((prevRange) => [prevRange[0], [rowIndex, colIndex]]);
    }
  };

  // Вычисление среднего значения
  const calculateAverage = () => {
    if (selectedRange.length === 2) {
      const [startRow, startCol] = selectedRange[0];
      const [endRow, endCol] = selectedRange[1];
      let sum = 0;
      let count = 0;

      for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
        for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
          const value = parseFloat(rows[r][c].value) || 0;
          sum += value;
          count++;
        }
      }
      return count ? (sum / count).toFixed(2) : 0;
    }
    return 0;
  };

  // Добавление диапазона в формулу
  const insertRangeIntoFormula = () => {
    if (selectedRange.length === 2) {
      const startCell = indexToCell(selectedRange[0][0], selectedRange[0][1]);
      const endCell = indexToCell(selectedRange[1][0], selectedRange[1][1]);
      const range = `${startCell}:${endCell}`;
      setFormula((prevFormula) => `${prevFormula}${range}`);
    }
  };

  // Сохранение таблицы как XLSX файл
  const saveAsXlsx = () => {
    const sheetData = rows.map((row) => row.map((cell) => cell.value));
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([excelBuffer], { type: "application/octet-stream" });
    saveAs(blob, "table.xlsx");
  };

  // Отображение выделенного диапазона
  const renderSelectedRangeLabel = () => {
    if (selectedRange.length === 2) {
      const startCell = indexToCell(selectedRange[0][0], selectedRange[0][1]);
      const endCell = indexToCell(selectedRange[1][0], selectedRange[1][1]);
      return <div className="range-label">Выбрано: {startCell} : {endCell}</div>;
    }
    return <div className="range-label">Выбрано: &nbsp;</div>;
  };

  // Обновление таблицы с функциональностью клика по ячейкам
  const renderTable = () => {
    const columns = Array.from({ length: rows[0].length }, (_, i) =>
      String.fromCharCode(65 + i)
    );

    return (
      <div className="excel-table" ref={tableRef}>
        <div className="row header-row">
          <div className="cell header-cell"></div>
          {columns.map((col, index) => (
            <div key={index} className="cell header-cell">
              {col}
            </div>
          ))}
        </div>

        {rows.map((row, rowIndex) => (
          <div key={rowIndex} className="row">
            <div className="cell header-cell">{rowIndex + 1}</div>
            {row.map((cell, colIndex) => (
              <input
                key={colIndex}
                className={`cell ${selectedCell?.row === rowIndex && selectedCell?.col === colIndex
                    ? "selected"
                    : selectedRange.length === 2 &&
                      rowIndex >= Math.min(selectedRange[0][0], selectedRange[1][0]) &&
                      rowIndex <= Math.max(selectedRange[0][0], selectedRange[1][0]) &&
                      colIndex >= Math.min(selectedRange[0][1], selectedRange[1][1]) &&
                      colIndex <= Math.max(selectedRange[0][1], selectedRange[1][1])
                    ? "range-selected"
                    : ""}`}
                value={cell.value}
                onChange={(e) => handleCellChange(rowIndex, colIndex, e.target.value)}
                onClick={() => handleCellClick(rowIndex, colIndex)}
                onMouseDown={() => startRangeSelection(rowIndex, colIndex)}
                onMouseEnter={(e) => handleMouseEnter(rowIndex, colIndex)}
                onMouseUp={handleMouseUp}
              />
            ))}
          </div>
        ))}
      </div>
    );
  };

  // Обработка импорта Excel файла
  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        setRows(sheetData.map(row => row.map(value => ({ value, style: {} }))));
      };
      reader.readAsArrayBuffer(file);
    }
  };

  return (
    <div className="app">
      <div className="toolbar">
        {renderSelectedRangeLabel()}
        <input
          id="formulaInput"
          type="text"
          placeholder="Введите значение или формулу"
          value={formula}
          onChange={handleFormulaChange}
          onKeyDown={handleKeyDown} // Применение формулы при нажатии Enter
        />
        <button onClick={insertRangeIntoFormula}>Вставить диапазон</button>
        <button onClick={() => alert(`Среднее значение: ${calculateAverage()}`)}>Среднее значение</button>
        <button onClick={saveAsXlsx}>Сохранить как .xlsx</button>
        <input type="file" accept=".xlsx, .xls" onChange={handleFileChange} />
      </div>

      {renderTable()}
    </div>
  );
};

export default App;
