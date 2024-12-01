import React, { useState, useEffect, useRef, useCallback } from "react";
import "./App.css";
import * as XLSX from "sheetjs-style"; // or "xlsx" depending on your preference
import { saveAs } from "file-saver";
import { evaluate } from "mathjs"; // Для вычислений формул
import { createTheme, ThemeProvider } from '@mui/material/styles';
import { AppBar, Toolbar, Typography, Container, Button, TextField, IconButton } from '@mui/material';
import { useDropzone } from 'react-dropzone';
import { BrowserRouter as Router, Routes, Route, useLocation } from 'react-router-dom';
import HomePage from './HomePage'; // Импорт нового компонента HomePage
import { copy, read } from 'clipboard-copy';

const theme = createTheme({
  palette: {
    primary: {
      main: '#10793F',
    },
  },
});

const App = () => {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<HomePage />} />
        <Route path="/app" element={<AppContent />} />
      </Routes>
    </Router>
  );
};

const AppContent = () => {
  // Создание пустой ячейки с начальными стилями
  const createEmptyCell = () => ({
    value: "",
    formula: "",
    style: {
      font: {
        bold: false,
        italic: false,
        name: "Arial",
        sz: 11
      }
    },
    type: "s"
  });

  const [rows, setRows] = useState(() => {
    return Array(20).fill().map(() => 
      Array(10).fill().map(() => createEmptyCell())
    );
  });
  const [selectedCell, setSelectedCell] = useState(null); // Хранение выделенной ячейки
  const [selectedRange, setSelectedRange] = useState([]); // Для хранения диапазона
  const [formula, setFormula] = useState(""); // Формула для ввода
  const [columnWidths, setColumnWidths] = useState(Array(100).fill(100)); // Ширина столбцов 
  const tableRef = useRef(null); // Ссылка на таблицу для контроля за координатами
  const [resizingColumn, setResizingColumn] = useState(null); // Индекс изменяемого столбца
  const [startX, setStartX] = useState(null); // Начальная позиция мыши для изменения размера
  const [fileName, setFileName] = useState('');
  const [isSelectingRange, setIsSelectingRange] = useState(false);
  const [copiedData, setCopiedData] = useState(null);
  const [ctrlPressed, setCtrlPressed] = useState(false); // Статус клавиши Ctrl
  const [isDarkMode, setIsDarkMode] = useState(false);
  const cellRefs = useRef([]); // Рефы для ячеек

  useEffect(() => {
    document.body.classList.toggle('dark-theme', isDarkMode);
    document.body.classList.toggle('light-theme', !isDarkMode);
  }, [isDarkMode]);

  const toggleTheme = () => {
    setIsDarkMode(prevMode => !prevMode);
  };

  const calculateFormula = (formula) => {
    if (!formula.startsWith('=')) return formula;
  
    try {
      // Убираем '=' в начале строки
      let calculatedValue = formula.slice(1);
  
      // Обрабатываем функции (например, SUM, AVERAGE)
      calculatedValue = handleFunctions(calculatedValue);
  
      // Замена ссылок на ячейки их значениями
      calculatedValue = calculatedValue.replace(/([A-Z])(\d+)/g, (match, col, row) => {
        const colIndex = col.charCodeAt(0) - 65; // Преобразуем букву колонки в индекс
        const rowIndex = parseInt(row, 10) - 1; // Преобразуем строку в индекс (1-базовый в 0-базовый)
        if (
          rowIndex >= 0 &&
          rowIndex < rows.length &&
          colIndex >= 0 &&
          colIndex < rows[0].length
        ) {
          const cellValue = rows[rowIndex][colIndex].value;
          return !isNaN(cellValue) ? cellValue : '0'; // Возвращаем значение или 0
        }
        return '0'; // Если ссылка некорректна, подставляем 0
      });
  
      // Вычисляем результат выражения
      const result = evaluate(calculatedValue);
      return result;
    } catch (error) {
      console.error("Ошибка в вычислении формулы:", formula, error);
      return "#ERROR"; // Возвращаем ошибку вместо формулы
    }
  };
  

  const updateFormulas = useCallback(() => {
    const newRows = [...rows];
    let hasChanges = false;
  
    // Проходим по всем строкам и ячейкам
    for (let row = 0; row < newRows.length; row++) {
      for (let col = 0; col < newRows[row].length; col++) {
        const cell = newRows[row][col];
        if (cell.formula && cell.formula.startsWith('=')) {
          try {
            // Пересчитываем формулу
            const newValue = calculateFormula(cell.formula);
            if (newValue !== cell.value) {
              hasChanges = true; // Отмечаем изменения
              newRows[row][col] = {
                ...cell,
                value: newValue,
                type: !isNaN(newValue) && newValue !== "" ? "n" : "s", // Обновляем тип
              };
            }
          } catch (error) {
            console.error(`Ошибка обновления формулы в ячейке (${row}, ${col}):`, error);
          }
        }
      }
    }
  
    // Если изменения были, обновляем состояние
    if (hasChanges) {
      setRows(newRows);
    }
  }, [rows]);
  

  const handleFileChange = useCallback((acceptedFiles) => {
    const file = acceptedFiles[0];
    setFileName(file.name);
    
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { 
        type: 'array',
        cellFormula: true,
        cellStyles: true,
        cellDates: true
      });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Размеры таблицы
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      const numRows = range.e.r + 1;
      const numCols = range.e.c + 1;
      
      // Создание новой таблицы
      const newRows = Array(numRows).fill().map(() => 
        Array(numCols).fill().map(() => createEmptyCell())
      );

      // Заполняем данными из файла Excel
      for (let row = 0; row < numRows; row++) {
        for (let col = 0; col < numCols; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell) {
            let value = cell.v;
            let type = cell.t;
            let formula = '';

            // Обработка формул
            if (cell.f) {
              formula = cell.f.startsWith('=') ? cell.f : '=' + cell.f;
              // Если есть формула, но нет вычисленного значения, вычисляем
              try {
                const calcValue = calculateFormula(formula);
                value = calcValue === "#ERROR!" ? value : calcValue;
              } catch (error) {
                console.error("Error calculating formula:", error);
              }
            }

            // Обработка чисел с запятыми
            if (!formula && typeof value === 'string' && value.includes(',')) {
              // Пробуем преобразовать строку с запятой в число
              const numberValue = parseFloat(value.replace(',', '.'));
              if (!isNaN(numberValue)) {
                value = numberValue;
                type = 'n';
              }
            }

            // n - число, s - строка, проверяем тип данных
            if (type === 'n' || (!isNaN(value) && value !== '')) {
              type = 'n';
            } else {
              type = 's';
            }

            newRows[row][col] = {
              ...createEmptyCell(),
              value: value,
              formula: formula,
              type: type,
              style: {
                ...createEmptyCell().style,
                bold: cell.s?.font?.bold || false,
                italic: cell.s?.font?.italic || false
              }
            };
          }
        }
      }
      
      setRows(newRows);
      // Обновляем формулы после загрузки файла
      setTimeout(updateFormulas, 0);
    };
    
    reader.readAsArrayBuffer(file);
  }, [calculateFormula, updateFormulas]);


  const { getRootProps, getInputProps } = useDropzone({
    onDrop: handleFileChange,
    multiple: false,
    accept: '.xlsx, .xlsm',
  });

  const handleCellChange = (row, col, value) => {
    const newRows = [...rows];
  
    if (value.startsWith('=')) {
      try {
        // Сохраняем формулу и вычисляем значение
        const result = calculateFormula(value);
        newRows[row][col] = {
          ...newRows[row][col],
          value: result,
          formula: value,
          type: !isNaN(result) ? "n" : "s",
        };
        setFormula(value);
      } catch (error) {
        // В случае ошибки сохраняем как текст
        newRows[row][col] = {
          ...newRows[row][col],
          value: value,
          formula: value,
          type: "s",
        };
        setFormula(value);
      }
    } else {
      const type = !isNaN(value) && value !== "" ? "n" : "s";
      newRows[row][col] = {
        ...newRows[row][col],
        value: value,
        formula: value,
        type: type,
      };
      setFormula(value);
    }
  
    setRows(newRows);
  
    setTimeout(updateFormulas, 0);
  };
  

  const deepCopy = (obj) => {
    if (typeof obj !== 'object' || obj === null) {
      return obj;
    }

    const copy = Array.isArray(obj) ? [] : {};
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        copy[key] = deepCopy(obj[key]);
      }
    }
    return copy;
  }
  const clearRange = () => {
    if (selectedRange.length === 2) {
      const newRows = [...rows];
      const [start, end] = selectedRange;
      const startRow = Math.min(start[0], end[0]);
      const endRow = Math.max(start[0], end[0]);
      const startCol = Math.min(start[1], end[1]);
      const endCol = Math.max(start[1], end[1]);
  
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          newRows[r][c] = {
            ...createEmptyCell(),
            style: newRows[r][c].style
          };
        }
      }
      setRows(newRows);
    }
  };

  // Функция копирования
  const copyRange = useCallback((selectedRange, rows) => {
    if (!selectedRange || selectedRange.length !== 2) return;
    
    const [start, end] = selectedRange;
    const startRow = Math.min(start[0], end[0]);
    const endRow = Math.max(start[0], end[0]);
    const startCol = Math.min(start[1], end[1]);
    const endCol = Math.max(start[1], end[1]);

    const copiedData = [];
    for (let row = startRow; row <= endRow; row++) {
      const rowData = [];
      for (let col = startCol; col <= endCol; col++) {
        rowData.push(rows[row][col].value);
      }
      copiedData.push(rowData);
    }

    const textToCopy = copiedData.map(row => row.join('\t')).join('\n');
    navigator.clipboard.writeText(textToCopy);
    setCopiedData(copiedData);
  }, []);

  // Функция вставки
  const pasteRangeValues = useCallback(async (selectedRange, rows) => {
    if (!selectedRange || selectedRange.length !== 2) return;

    try {
      const clipboardText = await navigator.clipboard.readText();
      const pasteData = clipboardText
        .split(/\r?\n/)
        .map(row => row.split('\t'))
        .filter(row => row.length > 0 && row.some(cell => cell !== ''));

      const [start, end] = selectedRange;
      const targetStartRow = Math.min(start[0], end[0]);
      const targetStartCol = Math.min(start[1], end[1]);

      const newRows = [...rows];

      pasteData.forEach((rowData, rowIndex) => {
        rowData.forEach((value, colIndex) => {
          const targetRow = targetStartRow + rowIndex;
          const targetCol = targetStartCol + colIndex;

          if (targetRow < newRows.length && targetCol < newRows[0].length) {
            newRows[targetRow][targetCol] = {
              ...createEmptyCell(),
              value: value,
              formula: value.toString().startsWith('=') ? value : '',
              type: !isNaN(value) && value !== "" ? "n" : "s"
            };
          }
        });
      });

      setRows(newRows);
      setTimeout(updateFormulas, 0);
    } catch (err) {
      alert('Не удалось вставить данные из буфера обмена');
    }
  }, [createEmptyCell, updateFormulas]);

  useEffect(() => {
    const handleKeyDown = (event) => {
      if (event.key === "Control" || event.key === "Meta") {
        setCtrlPressed(true);
      }
      if (event.key === "Delete") {
        if (event.ctrlKey || event.metaKey) {
          // Ctrl+Delete - очистка диапазона
          event.preventDefault();
          clearRange();
          return;
        } else {
          // Обычный Delete - очистка значения
          const newRows = [...rows];
          if (selectedRange.length === 2) {
            const [startRow, startCol] = selectedRange[0];
            const [endRow, endCol] = selectedRange[1];
            for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
              for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
                newRows[r][c].value = "";
                newRows[r][c].formula = "";
              }
            }
          } else if (selectedCell) {
            newRows[selectedCell.row][selectedCell.col].value = "";
            newRows[selectedCell.row][selectedCell.col].formula = "";
          }
          setRows(newRows);
        }
      } if (selectedCell && !isSelectingRange) {
        let newRow = selectedCell.row;
        let newCol = selectedCell.col;
  
        switch (event.key) {
          case 'ArrowUp':
            newRow = Math.max(selectedCell.row - 1, 0);
            break;
          case 'ArrowDown':
            newRow = Math.min(selectedCell.row + 1, rows.length - 1);
            break;
          case 'ArrowLeft':
            newCol = Math.max(selectedCell.col - 1, 0);
            break;
          case 'ArrowRight':
            newCol = Math.min(selectedCell.col + 1, rows[0].length - 1);
            break;
          default:{
            // Обработка Copy/Paste
            if (event.ctrlKey || event.metaKey) {
              const key = event.key.toLowerCase();
              // Проверяем все возможные варианты клавиш для копирования
              if (key === 'c' || key === 'с') { // латинская и русская 'c'
                event.preventDefault();
                if (selectedRange.length === 2) {
                  copyRange(selectedRange, rows);
                } else if (selectedCell) {
                  copyRange(
                    [[selectedCell.row, selectedCell.col], 
                     [selectedCell.row, selectedCell.col]], 
                    rows
                  );
                }
                return;
              }
              // Проверяем все возможные варианты клавиш для вставки
              if (key === 'v' || key === 'м') { // латинская 'v' и русская 'м'
                event.preventDefault();
                if (selectedCell) {
                  const startRange = [selectedCell.row, selectedCell.col];
                  const endRange = [selectedCell.row, selectedCell.col];
                  pasteRangeValues([startRange, endRange], rows);
                }
                return;
              }
            }
            return;
          }
        }
  
        // Обновляем выбранную ячейку и формулу
        const newCell = rows[newRow][newCol];
        setSelectedCell({ row: newRow, col: newCol });
        setSelectedRange([]);
        setFormula(newCell.formula || newCell.value.toString());

        // Предотвращаем стандартное поведение клавиш стрелок
        if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) {
          event.preventDefault();
        }
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
  }, [selectedCell, selectedRange, rows, copyRange, pasteRangeValues, clearRange]);


  // Преобразуем индексы строки и столбца в ссылку на ячейку
  const indexToCell = (row, col) => {
    const colName = String.fromCharCode(65 + col); // Преобразуем индекс колонки в букву
    return `${colName}${row + 1}`;
  };

  const applyFormula = () => {
    if (!selectedCell) return;

    const { row, col } = selectedCell;
    handleCellChange(row, col, formula);
  };

  const handleFormulaInputChange = (e) => {
    const newValue = e.target.value;
    setFormula(newValue);
    
    // Если есть выбранная ячейка, обновляем её формулу
    if (selectedCell) {
      const newRows = [...rows];
      newRows[selectedCell.row][selectedCell.col].formula = newValue;
      setRows(newRows);
    }
  };

  const handleFormulaKeyDown = (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      applyFormula();
      e.target.blur();
    }
  };

  const startRangeSelection = (rowIndex, colIndex, event) => {
    const formulaInput = document.getElementById('formulaInput');
    const isFormulaActive = formulaInput === document.activeElement && formula.startsWith('=');
  
    if (isFormulaActive) {
      event.preventDefault();
      event.stopPropagation();
      const colLetter = String.fromCharCode(65 + colIndex);
      const cellRef = `${colLetter}${rowIndex + 1}`;
      const cursorPos = formulaInput.selectionStart;
      const newFormula = formula.slice(0, cursorPos) + cellRef + formula.slice(cursorPos);
      setFormula(newFormula);
      formulaInput.focus();
      formulaInput.setSelectionRange(cursorPos + cellRef.length, cursorPos + cellRef.length);
      return;
    }
  
    if (event.shiftKey && selectedCell) {
      setSelectedRange([
        [selectedCell.row, selectedCell.col],
        [rowIndex, colIndex]
      ]);
    } else {
      setSelectedCell({ row: rowIndex, col: colIndex });
      setSelectedRange([]);
  
      const cell = rows[rowIndex][colIndex];
  
      if (cell) {
        if (cell.formula) {
          setFormula(cell.formula);
        } else {
          setFormula(cell.value ? cell.value.toString() : ""); // Учитываем, что value может быть undefined
        }
      } else {
        setFormula(""); // Если ячейка полностью отсутствует
      }
    }
  
    if (!event.shiftKey) {
      setIsSelectingRange(true);
      setSelectedRange([[rowIndex, colIndex], [rowIndex, colIndex]]);
    }
  };
  

  // Обработка мыши для выделения диапазона
  const handleMouseMove = (e, rowIndex, colIndex) => {
    if (isSelectingRange) {
      const startRow = Math.min(selectedRange[0][0], rowIndex);
      const startCol = Math.min(selectedRange[0][1], colIndex);
      const endRow = Math.max(selectedRange[0][0], rowIndex);
      const endCol = Math.max(selectedRange[0][1], colIndex);
      setSelectedRange([[startRow, startCol], [endRow, endCol]]);
    }
  };

  // Завершение выделения диапазона
  const handleMouseUp = () => {
    setIsSelectingRange(false);
    if (selectedRange.length === 2) {
      const [start, end] = selectedRange;
      const startRow = Math.min(start[0], end[0]);
      const startCol = Math.min(start[1], end[1]);
      const endRow = Math.max(start[0], end[0]);
      const endCol = Math.max(start[1], end[1]);
      setSelectedRange([[startRow, startCol], [endRow, endCol]]);
    }
  };

  const handleCellClick = (row, col, event) => {
    // удалить...
  };

  // Выбор диапазона без потери фокуса
  const handleMouseEnter = (rowIndex, colIndex) => {
    if (isSelectingRange) {
      handleMouseMove(null, rowIndex, colIndex);
    }
  };

  const insertSingleCellIntoFormula = () => {
    if (!selectedCell && (!selectedRange || selectedRange.length !== 2)) return;

    let cellRef = '';
    
    if (selectedCell) {
      // Одиночная выбранная ячейка
      const colLetter = String.fromCharCode(65 + selectedCell.col);
      cellRef = `${colLetter}${selectedCell.row + 1}`;
    } else {
      // Проверяем, является ли диапазон одной ячейкой
      const [start, end] = selectedRange;
      if (start[0] === end[0] && start[1] === end[1]) {
        const colLetter = String.fromCharCode(65 + start[1]);
        cellRef = `${colLetter}${start[0] + 1}`;
      } else {
        return; // Если выбрано больше одной ячейки, не делаем ничего
      }
    }

    const formulaInput = document.getElementById('formulaInput');
    if (formulaInput) {
      const cursorPos = formulaInput.selectionStart;
      const currentFormula = formula;
      const newFormula = currentFormula.slice(0, cursorPos) + cellRef + currentFormula.slice(cursorPos);
      setFormula(newFormula);
      
      // Устанавливаем курсор после вставленной ссылки
      setTimeout(() => {
        formulaInput.focus();
        formulaInput.setSelectionRange(cursorPos + cellRef.length, cursorPos + cellRef.length);
      }, 0);
    }
  };

    const saveAsXlsx = (format = "xlsx") => {
      const sheetData = rows.map(row =>
        row.map(cell => {
          const cellObj = {
            t: cell.type, // Тип ячейки: 'n' для чисел, 's' для строк
            v: cell.value // Значение ячейки
          };
    
          // Добавляем формулы, если есть
          if (cell.formula && cell.formula.startsWith('=')) {
            // Для всех форматов оставляем формулы
            cellObj.f = cell.formula.substring(1); // Убираем '=' для Excel
          }
    
          // Добавляем стили, если они есть
          if (cell.style) {
            cellObj.s = {};
            if (cell.style.bold || cell.style.italic) {
              cellObj.s.font = {
                bold: cell.style.bold || false,
                italic: cell.style.italic || false
              };
            }
            if (cell.style.fill) {
              cellObj.s.fill = { fgColor: { rgb: cell.style.fill } };
            }
          }
    
          return cellObj;
        })
      );
    
      const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    
      // Устанавливаем ширину столбцов, если определена
      if (typeof columnWidths !== "undefined") {
        worksheet['!cols'] = columnWidths.map(width => ({ wpx: width }));
      }
    
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    
      // Определяем параметры для формата
      let bookType = "xlsx";
      let mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      let fileName = "Лист.xlsx";
    
      switch (format) {
        case "xls":
          bookType = "xls";
          mimeType = "application/vnd.ms-excel";
          fileName = "Лист.xls";
          break;
        case "xlsm":
          bookType = "xlsm";
          mimeType = "application/vnd.ms-excel.sheet.macroEnabled.12";
          fileName = "Лист.xlsm";
          break;
        default:
          break;
      }
    
      // Буфер для файла
      const excelBuffer = XLSX.write(workbook, { bookType, type: "array", cellFormula: true });
    
      // Сохраняем файл
      const blob = new Blob([excelBuffer], { type: mimeType });
      saveAs(blob, fileName);
    };
    
    // Функция для сохранения в разные форматы
    const handleSaveAs = () => {
      const userInput = prompt("Введите формат для файла (xlsx, xlsm):");

      // Проверяем, что пользователь не нажал "Отмена"
      if (userInput === null) {
        return; // Просто выходим из функции, если пользователь отменил ввод
      }

      const format = userInput.toLowerCase();
      if (["xlsx", "xlsm"].includes(format)) {
        saveAsXlsx(format);
      } else {
        alert("Неверный формат. Пожалуйста, введите 'xlsx' или 'xlsm'.");
      }
    };

  // Функция для изменения размера столбца
  const handleColumnResizeStart = (columnIndex, e) => {
    setResizingColumn(columnIndex);
    setStartX(e.clientX);
  };

  // Отображение выделенного диапазона
  const renderSelectedRangeLabel = () => {
    if (selectedRange.length === 2) {
      const startCell = indexToCell(selectedRange[0][0], selectedRange[0][1]);
      const endCell = indexToCell(selectedRange[1][0], selectedRange[1][1]);
      return <div className="range-label">Выбрано: {startCell}:{endCell}</div>;
    }
    return <div className="range-label">Выбрано: &nbsp;</div>;
  };

  const handleColumnResize = (e) => {
    if (resizingColumn !== null && startX !== null) {
      const newWidths = [...columnWidths];
      const deltaX = e.clientX - startX;
      newWidths[resizingColumn] = Math.max(50, columnWidths[resizingColumn] + deltaX);
      setColumnWidths(newWidths);
      setStartX(e.clientX);
    }
  };

  const handleColumnResizeEnd = () => {
    setResizingColumn(null);
    setStartX(null);
  };

  const handleMouseDown = (index) => (e) => {
    const startX = e.clientX;
    const startWidth = parseInt(document.defaultView.getComputedStyle(e.target.parentNode).width, 10);

    const doDrag = (e) => {
      const newWidth = startWidth + e.clientX - startX;
      const newRows = rows.map(row =>
        row.map((cell, i) =>
          i === index ? { ...cell, style: { ...cell.style, width: `${newWidth}px` } } : cell
        )
      );
      setRows(newRows);
    };

    const stopDrag = () => {
      document.removeEventListener('mousemove', doDrag);
      document.removeEventListener('mouseup', stopDrag);
    };

    document.addEventListener('mousemove', doDrag);
    document.addEventListener('mouseup', stopDrag);
  };

  // Обновление таблицы с функциональностью клика по ячейкам
  const renderTable = () => {
    const columns = Array.from({ length: rows[0].length }, (_, i) =>
      String.fromCharCode(65 + i)
    );
    setTimeout(updateFormulas, 0);
    return (
      <div
        className="excel-table"
        ref={tableRef}
        onMouseMove={handleColumnResize}
        onMouseUp={handleColumnResizeEnd}
        style={{ width: '100%', overflowX: 'auto' }}
      >
        <div className="row header-row">
          <div className="cell header-cell"></div>
          {columns.map((col, index) => (
            <div
              key={index}
              className="cell header-cell"
              style={{ width: columnWidths[index], position: 'relative' }}
            >
              {col}
              <div
                className="column-resizer"
                onMouseDown={(e) => handleColumnResizeStart(index, e)}
              ></div>
            </div>
          ))}
        </div>

        {rows.map((row, rowIndex) => (
          <div key={rowIndex} className="row">
            <div className="cell header-cell">{rowIndex + 1}</div>
            {row.map((cell, colIndex) => (
              <div
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
                style={{ width: columnWidths[colIndex] }}
                onMouseDown={(e) => startRangeSelection(rowIndex, colIndex, e)}
                onMouseEnter={(e) => handleMouseEnter(rowIndex, colIndex)}
                onMouseUp={handleMouseUp}
              >
                <input
                  ref={(el) => {
                    if (!cellRefs.current[rowIndex]) {
                      cellRefs.current[rowIndex] = [];
                    }
                    cellRefs.current[rowIndex][colIndex] = el;
                  }}
                
                  type="text"
                  className="cell-input"
                  value={selectedCell?.row === rowIndex && selectedCell?.col === colIndex ? cell.value : cell.value}
                  onChange={(e) => handleCellChange(rowIndex, colIndex, e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      handleCellChange(rowIndex, colIndex, e.target.value);
                    }
                  }}
                  onClick={(e) => handleCellClick(rowIndex, colIndex, e)}
                  style={{
                    fontWeight: cell.style?.font?.bold ? 'bold' : 'normal',
                    fontStyle: cell.style?.font?.italic ? 'italic' : 'normal'
                  }}
                />
              </div>
            ))}
          </div>
        ))}
      </div>
    );
  };

  const addRow = (index) => {
    const newRows = [...rows];
    const numCols = newRows[0].length;
    const newRow = Array(numCols).fill(createEmptyCell());
    newRows.splice(index, 0, newRow);
    setRows(newRows);
  };

  const addColumn = (index) => {
    const newRows = rows.map((row) => {
      const newRow = [...row];
      newRow.splice(index, 0, createEmptyCell());
      return newRow;
    });
    setRows(newRows);
    const newColumnWidths = [...columnWidths];
    newColumnWidths.splice(index, 0, 100); // Default width for new column
    setColumnWidths(newColumnWidths);
  };

  const handleAddRow = () => {
    addRow(rows.length);
  };

  const handleAddColumn = () => {
    addColumn(rows[0].length);
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
          const value = parseFloat(rows[r][c].value);
          if (!isNaN(value)) {
            sum += value;
            count++;
          }
        }
      }
      return count ? (sum / count).toFixed(2) : 0;
    }
    return 0;
  };

  // Преобразование столбца (буквы) в индекс
  const colToIndex = (col) => {
    let index = 0;
    col = col.toUpperCase();  // к верхнему регистру
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.charCodeAt(i) - 65 + 1); // 'A' = 65, начинаем с 1
    }
    return index - 1;  // Корректируем индексы (т.к. 0)
  };

  // Обработка функций в формулах
  const handleFunctions = (formula) => {
    const sumRegex = /(?:СУММ|SUM)\((([A-Z])(\d+):([A-Z])(\d+))\)/i;
    const avgRegex = /(?:СРЗНАЧ|AVG)\((([A-Z])(\d+):([A-Z])(\d+))\)/i;
    const sumifRegex = /(?:СУММЕСЛИ|SUMIF)\((([A-Z])(\d+):([A-Z])(\d+)),([^)]+)\)/i;
    const minRegex = /(?:МИН|MIN)\((([A-Z])(\d+):([A-Z])(\d+))\)/i;
    const maxRegex = /(?:МАКС|MAX)\((([A-Z])(\d+):([A-Z])(\d+))\)/i;
    const countRegex = /(?:СЧЕТ|COUNT)\((([A-Z])(\d+):([A-Z])(\d+))\)/i;
    const countifRegex = /(?:СЧЕТЕСЛИ|COUNTIF)\((([A-Z])(\d+):([A-Z])(\d+)),([^)]+)\)/i;

    // Функция СУММ/SUM
    if (sumRegex.test(formula)) {
      const match = formula.match(sumRegex);
      const startCol = match[2];
      const startRow = parseInt(match[3]) - 1;
      const endCol = match[4];
      const endRow = parseInt(match[5]) - 1;

      let sum = 0;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol.charCodeAt(0) - 65; col <= endCol.charCodeAt(0) - 65; col++) {
          const value = rows[row][col].value;
          if (!isNaN(value)) {
            sum += parseFloat(value);
          }
        }
      }
      return formula.replace(sumRegex, sum);
    } 

    // Функция AVERAGE
    else if (avgRegex.test(formula)) {
      const match = formula.match(avgRegex);
      const startCol = match[2];
      const startRow = parseInt(match[3]) - 1;
      const endCol = match[4];
      const endRow = parseInt(match[5]) - 1;

      let sum = 0;
      let count = 0;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol.charCodeAt(0) - 65; col <= endCol.charCodeAt(0) - 65; col++) {
          const value = rows[row][col].value;
          if (!isNaN(value)) {
            sum += parseFloat(value);
            count++;
          }
        }
      }
      return formula.replace(avgRegex, (count > 0 ? sum / count : 0));
    }

    // Функция СУММЕСЛИ/SUMIF
    else if (sumifRegex.test(formula)) {
      const match = formula.match(sumifRegex);
      const startCol = match[2];
      const startRow = parseInt(match[3]) - 1;
      const endCol = match[4];
      const endRow = parseInt(match[5]) - 1;
      let criteria = match[6].trim();

      // Удаляем кавычки из критерия если они есть
      if (criteria.startsWith('"') && criteria.endsWith('"')) {
        criteria = criteria.slice(1, -1);
      }

      let sum = 0;
      for (let row = startRow; row <= endRow; row++) {
        const col = startCol.charCodeAt(0) - 65;
        const value = rows[row][col].value;
        
        // Проверяем соответствие критерию
        let meetsCondition = false;
        if (criteria.startsWith('>')) {
          meetsCondition = parseFloat(value) > parseFloat(criteria.slice(1));
        } else if (criteria.startsWith('<')) {
          meetsCondition = parseFloat(value) < parseFloat(criteria.slice(1));
        } else if (criteria.startsWith('>=')) {
          meetsCondition = parseFloat(value) >= parseFloat(criteria.slice(2));
        } else if (criteria.startsWith('<=')) {
          meetsCondition = parseFloat(value) <= parseFloat(criteria.slice(2));
        } else if (criteria.startsWith('<>')) {
          meetsCondition = value.toString() !== criteria.slice(2);
        } else if (criteria.startsWith('=')) {
          meetsCondition = value.toString() === criteria.slice(1);
        } else {
          // Прямое сравнение
          meetsCondition = value.toString() === criteria;
        }

        // Если условие выполняется и значение числовое, добавляем к сумме
        if (meetsCondition && !isNaN(value)) {
          sum += parseFloat(value);
        }
      }
      return formula.replace(sumifRegex, sum);
    }

    // Функция MIN
    else if (minRegex.test(formula)) {
      const match = formula.match(minRegex);
      const startCol = match[2];
      const startRow = parseInt(match[3]) - 1;
      const endCol = match[4];
      const endRow = parseInt(match[5]) - 1;

      let minValue = Infinity;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol.charCodeAt(0) - 65; col <= endCol.charCodeAt(0) - 65; col++) {
          const value = parseFloat(rows[row][col].value);
          if (!isNaN(value) && value < minValue) {
            minValue = value;
          }
        }
      }
      return formula.replace(minRegex, minValue === Infinity ? 0 : minValue);
    }

    // Функция MAX
    else if (maxRegex.test(formula)) {
      const match = formula.match(maxRegex);
      const startCol = match[2];
      const startRow = parseInt(match[3]) - 1;
      const endCol = match[4];
      const endRow = parseInt(match[5]) - 1;

      let maxValue = -Infinity;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol.charCodeAt(0) - 65; col <= endCol.charCodeAt(0) - 65; col++) {
          const value = parseFloat(rows[row][col].value);
          if (!isNaN(value) && value > maxValue) {
            maxValue = value;
          }
        }
      }
      return formula.replace(maxRegex, maxValue === -Infinity ? 0 : maxValue);
    }

    // Функция COUNT/СЧЕТ
    else if (countRegex.test(formula)) {
      const match = formula.match(countRegex);
      const startCol = match[2];
      const startRow = parseInt(match[3]) - 1;
      const endCol = match[4];
      const endRow = parseInt(match[5]) - 1;

      let count = 0;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol.charCodeAt(0) - 65; col <= endCol.charCodeAt(0) - 65; col++) {
          const value = rows[row][col]?.value;
          if (!isNaN(value)) {
            count++;
          }
        }
      }
      return formula.replace(countRegex, count);
    }

    // Функция COUNTIF/СЧЕТЕСЛИ
    else if (countifRegex.test(formula)) {
      const match = formula.match(countifRegex);
      const startCol = match[2];
      const startRow = parseInt(match[3]) - 1;
      const endCol = match[4];
      const endRow = parseInt(match[5]) - 1;
      let criteria = match[6].trim();

      // Удаляем кавычки из критерия если они есть
      if (criteria.startsWith('"') && criteria.endsWith('"')) {
        criteria = criteria.slice(1, -1);
      }

      let count = 0;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol.charCodeAt(0) - 65; col <= endCol.charCodeAt(0) - 65; col++) {
          const value = rows[row][col]?.value;

          // Проверяем соответствие критерию
          let meetsCondition = false;
          if (criteria.startsWith('>')) {
            meetsCondition = parseFloat(value) > parseFloat(criteria.slice(1));
          } else if (criteria.startsWith('<')) {
            meetsCondition = parseFloat(value) < parseFloat(criteria.slice(1));
          } else if (criteria.startsWith('>=')) {
            meetsCondition = parseFloat(value) >= parseFloat(criteria.slice(2));
          } else if (criteria.startsWith('<=')) {
            meetsCondition = parseFloat(value) <= parseFloat(criteria.slice(2));
          } else if (criteria.startsWith('<>')) {
            meetsCondition = value.toString() !== criteria.slice(2);
          } else if (criteria.startsWith('=')) {
            meetsCondition = value.toString() === criteria.slice(1);
          } else {
            // Прямое сравнение
            meetsCondition = value.toString() === criteria;
          }

          // Если условие выполняется, увеличиваем счётчик
          if (meetsCondition) {
            count++;
          }
        }
      }
      return formula.replace(countifRegex, count);
    }

    return formula;
  };
// плавающая верхняя панель 
  return (
    <ThemeProvider theme={theme}>
      <div className={`app ${isDarkMode ? 'dark-theme' : 'light-theme'}`}>
        <AppBar position="sticky"> 
          <Toolbar>
            <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
              Zxcel
            </Typography>
            <Button color="inherit" onClick={saveAsXlsx} sx={{ marginRight: 2 }}>Сохранить как .xlsx</Button>
            <Button color="inherit" onClick={handleSaveAs} sx={{ marginRight: 2 }}>Сохранить как...</Button>
            <Button color="inherit" onClick={toggleTheme} sx={{ marginRight: 2 }}>
              {isDarkMode ? 'Светлая тема' : 'Тёмная тема'}
            </Button>
          </Toolbar>
        </AppBar>
        <Container maxWidth={false} sx={{ padding: 0, flex: 1, display: 'flex', flexDirection: 'column' }}>
          <div {...getRootProps()} style={{ border: '2px dashed #ccc', padding: '20px', textAlign: 'center', marginBottom: '20px' }}>
            <input {...getInputProps()} />
            <p>{fileName || 'Добавить файл'}</p>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', marginBottom: '20px', padding: '0 20px' }}>
            <TextField
              id="formulaInput"
              value={formula}
              onChange={handleFormulaInputChange}
              onKeyDown={handleFormulaKeyDown}
              variant="outlined"
              placeholder="Значение или формула"
              size="small"
              sx={{ width: '300px', marginRight: 2 }}
            />
            <Button variant="contained" color="primary" onClick={() => alert(`Среднее значение: ${calculateAverage()}`)} sx={{ marginRight: 2 }}>Среднее значение</Button>
            <Button variant="contained" color="primary" onClick={handleAddRow} sx={{ marginRight: 2 }}>Добавить строку</Button>
            <Button variant="contained" color="primary" onClick={handleAddColumn}>Добавить столбец</Button>
            <Button color="inherit" onClick={() => pasteRangeValues(selectedRange, rows)}>
              Вставить
            </Button>
            <Button color="inherit" onClick={clearRange }>
              Очистить
            </Button>

          </div>
          {renderSelectedRangeLabel()}
          <div className="table-container">
            {renderTable()}
          </div>
        </Container>
      </div>
    </ThemeProvider>
  );
};

export default App;
