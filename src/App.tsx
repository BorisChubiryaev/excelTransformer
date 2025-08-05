import React, { useState, useEffect } from "react";
import * as ExcelJS from "exceljs";

const App: React.FC = () => {
  const [firstFile, setFirstFile] = useState<File | null>(null);
  const [secondFile, setSecondFile] = useState<File | null>(null);
  const [diffData, setDiffData] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [checkedRows, setCheckedRows] = useState<Set<string>>(new Set());
  const [fullscreen, setFullscreen] = useState(false);

  const toggleFullscreen = () => {
    setFullscreen((prev) => !prev);
  };

  useEffect(() => {
    const handleEsc = (e: KeyboardEvent) => {
      if (e.key === "Escape" && fullscreen) {
        setFullscreen(false);
      }
    };
    window.addEventListener("keydown", handleEsc);
    return () => window.removeEventListener("keydown", handleEsc);
  }, [fullscreen]);

  const columnLabels: Record<string, string> = {
    "Column 4": "Адрес",
    "Column 13": "Этаж",
    "Column 20": "РМ",
    "Column 25": "Тип РМ",
    "Column 39": "Признак",
    "Column 40": "Таб. №",
    "Column 41": "ФИО",
    "Column 45": "Департамент",
    "Column 49": "Ответственный",
    "Column 52": "ДП",
    "Column 54": "Трайб",
    "Column 59": "Дата с",
    "Column 62": "Статус",
    "Column 64": "Кол-во",
  };

  const displayColumns = [
    "Checked",
    "Status",
    ...Object.keys(columnLabels),
    "ChangeType",
  ];

  const requiredColumns = Object.keys(columnLabels);

  const [filters, setFilters] = useState({
    status: new Set<string>(["БЫЛО", "СТАЛО", "НОВАЯ", "УДАЛЕНА"]),
    address: new Set<string>(),
    floor: new Set<string>(),
    city: new Set<string>(),
    quantity: new Set<string>(),
    changeType: new Set<string>(),
    toReserve: false,
    toPartner: false,
  });

  // === Извлечение города из адреса ===
  const extractCity = (address: string): string | null => {
    if (!address) return null;
    const match = address.match(/г\s+([А-Яа-яЁё-]+(?:\s[А-Яа-яЁё-]+)*)/);
    return match ? match[1].trim() : null;
  };

  const handleFileChange = (
    event: React.ChangeEvent<HTMLInputElement>,
    setFile: React.Dispatch<React.SetStateAction<File | null>>,
  ) => {
    if (event.target.files && event.target.files.length > 0) {
      setFile(event.target.files[0]);
    }
  };

  const filterColumns = (data: any[]) => {
    return data.map((row) => {
      const filteredRow: any = {};
      requiredColumns.forEach((column) => {
        filteredRow[column] = row[column] !== undefined ? row[column] : null;
      });
      return filteredRow;
    });
  };

  const readExcel = async (file: File) => {
    const workbook = new ExcelJS.Workbook();
    const arrayBuffer = await file.arrayBuffer();
    await workbook.xlsx.load(arrayBuffer);

    const worksheet = workbook.worksheets[2]; // Третий лист
    const jsonData: any[] = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Пропускаем заголовок

      const rowData: any = {};
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const colKey = `Column ${colNumber}`;
        if (requiredColumns.includes(colKey)) {
          rowData[colKey] = cell.value;
        }
      });

      // Добавляем строку, только если есть данные
      if (Object.values(rowData).some((val) => val != null)) {
        jsonData.push(rowData);
      }
    });

    if (jsonData.length > 0) {
      jsonData.pop();
    }

    return filterColumns(jsonData);
  };

  const compareData = (oldData: any[], newData: any[]) => {
    const keyColumn = "Column 20";
    const mapOldData = new Map(oldData.map((item) => [item[keyColumn], item]));
    const mapNewData = new Map(newData.map((item) => [item[keyColumn], item]));

    const diff: any[] = [];

    for (const [key, newItem] of mapNewData) {
      const oldItem = mapOldData.get(key);
      if (!oldItem) {
        diff.push({ type: "new", rm: key, old: null, new: newItem });
      } else if (JSON.stringify(oldItem) !== JSON.stringify(newItem)) {
        diff.push({ type: "changed", rm: key, old: oldItem, new: newItem });
      }
    }

    for (const [key, oldItem] of mapOldData) {
      if (!mapNewData.has(key)) {
        diff.push({ type: "deleted", rm: key, old: oldItem, new: null });
      }
    }

    return diff;
  };

  const handleUpload = async () => {
    if (!firstFile || !secondFile) {
      alert("Пожалуйста, загрузите оба файла.");
      return;
    }

    setLoading(true);
    try {
      const oldData = await readExcel(firstFile);
      const newData = await readExcel(secondFile);
      const differences = compareData(oldData, newData);
      setDiffData(differences);
    } catch (error) {
      console.error("Ошибка при обработке файлов:", error);
      alert(
        "Ошибка при чтении файлов. Проверьте, что это корректные .xlsx-файлы.",
      );
    } finally {
      setLoading(false);
    }
  };

  const renderValue = (value: any) => {
    if (
      value == null ||
      value === "" ||
      value === "null" ||
      value === "undefined"
    )
      return <span className="text-gray-400">—</span>;
    return value;
  };

  // === Подготовка строк с мета-информацией ===
  const allRows = diffData.flatMap((diff) => {
    const rows = [];

    if (diff.type === "changed") {
      const oldCity = extractCity(diff.old["Column 4"]);
      const newCity = extractCity(diff.new["Column 4"]);
      const oldQty = diff.old["Column 64"];
      const newQty = diff.new["Column 64"];

      rows.push({
        type: "old",
        data: diff.old,
        rm: diff.rm,
        city: oldCity,
        quantity: oldQty,
        source: "changed",
      });
      rows.push({
        type: "new",
        data: diff.new,
        rm: diff.rm,
        city: newCity,
        quantity: newQty,
        source: "changed",
      });
    } else if (diff.type === "new") {
      const city = extractCity(diff.new["Column 4"]);
      const qty = diff.new["Column 64"];
      rows.push({
        type: "new",
        data: diff.new,
        rm: diff.rm,
        city,
        quantity: qty,
        source: "new",
      });
    } else if (diff.type === "deleted") {
      const city = extractCity(diff.old["Column 4"]);
      const qty = diff.old["Column 64"];
      rows.push({
        type: "deleted",
        data: diff.old,
        rm: diff.rm,
        city,
        quantity: qty,
        source: "deleted",
      });
    }

    return rows;
  });

  // === Уникальные значения для фильтров ===
  const uniqueAddresses = Array.from(
    new Set(allRows.map((r) => r.data["Column 4"]).filter(Boolean)),
  ).sort();

  const uniqueFloors = Array.from(
    new Set(allRows.map((r) => r.data["Column 13"]).filter(Boolean)),
  ).sort();

  const uniqueCities = Array.from(
    new Set(allRows.map((r) => r.city).filter(Boolean)),
  ).sort();

  const uniqueQuantities = Array.from(
    new Set(
      allRows
        .map((r) => String(r.quantity))
        .filter((q) => q !== "null" && q !== "undefined" && q !== ""),
    ),
  ).sort((a, b) => Number(a) - Number(b));

  // Собираем все типы изменений (названия полей)
  const uniqueChangeTypes: string[] = [];

  // Добавляем типы из изменённых строк
  diffData.forEach((diff) => {
    if (diff.type === "changed") {
      const oldData = diff.old;
      const newData = diff.new;
      const changedFields = requiredColumns.filter(
        (key) => JSON.stringify(oldData[key]) !== JSON.stringify(newData[key]),
      );
      changedFields.forEach((key) => {
        const label = columnLabels[key];
        if (!uniqueChangeTypes.includes(label)) {
          uniqueChangeTypes.push(label);
        }
      });
    }
  });

  // Добавляем специальные типы
  ["Новая запись", "Удалена"].forEach((t) => {
    if (!uniqueChangeTypes.includes(t)) {
      uniqueChangeTypes.push(t);
    }
  });

  // Сортировка: "Признак" — первым
  uniqueChangeTypes.sort((a, b) => {
    if (a === "Признак") return -1;
    if (b === "Признак") return 1;
    return a.localeCompare(b);
  });

  // === Фильтрация ===
  const filteredRows = allRows.filter((row) => {
    const { data, type, city, quantity, source } = row;

    // Определяем отображаемый статус для фильтрации
    let filterStatus = "";
    if (source === "new") {
      filterStatus = "НОВАЯ";
    } else if (source === "deleted") {
      filterStatus = "УДАЛЕНА";
    } else {
      filterStatus = { old: "БЫЛО", new: "СТАЛО", deleted: "УДАЛЕНА" }[type];
    }

    if (!filters.status.has(filterStatus)) return false;

    if (filters.address.size > 0 && data["Column 4"]) {
      if (!filters.address.has(String(data["Column 4"]))) return false;
    }

    if (filters.floor.size > 0 && data["Column 13"]) {
      if (!filters.floor.has(String(data["Column 13"]))) return false;
    }

    if (filters.city.size > 0 && city) {
      if (!filters.city.has(city)) return false;
    }

    if (filters.quantity.size > 0) {
      const qtyStr = String(quantity);
      if (!filters.quantity.has(qtyStr)) return false;
    }

    // === Фильтр по типу изменения ===
    if (filters.changeType.size > 0) {
      let matchesChangeType = false;

      if (source === "new") {
        matchesChangeType = filters.changeType.has("Новая запись");
      } else if (source === "deleted") {
        matchesChangeType = filters.changeType.has("Удалена");
      } else if (type === "new") {
        // Это "СТАЛО" — ищем изменения
        const oldRow = allRows.find((r) => r.type === "old" && r.rm === row.rm);
        if (oldRow) {
          const changedFields = requiredColumns.filter((key) => {
            return (
              JSON.stringify(oldRow.data[key]) !== JSON.stringify(row.data[key])
            );
          });

          const changedLabels = changedFields.map((key) => columnLabels[key]);
          matchesChangeType = changedLabels.some((label) =>
            filters.changeType.has(label),
          );
        }
      }

      if (!matchesChangeType) return false;
    }

    // === Фильтр: Признак → Резерв ===
    if (filters.toReserve) {
      let matches = false;
      if (type === "new" && source !== "new" && source !== "deleted") {
        const oldRow = allRows.find((r) => r.type === "old" && r.rm === row.rm);
        if (oldRow) {
          const oldVal = oldRow.data["Column 39"];
          const newVal = row.data["Column 39"];
          // Проверяем, что стало "Резерв", а было что-то другое
          if (
            newVal &&
            String(newVal).trim() === "Резерв" &&
            (oldVal === null || String(oldVal).trim() !== "Резерв")
          ) {
            matches = true;
          }
        }
      }
      if (!matches) return false;
    }

    // === Фильтр: Признак → Партнер ===
    if (filters.toPartner) {
      let matches = false;
      if (type === "new" && source !== "new" && source !== "deleted") {
        const oldRow = allRows.find((r) => r.type === "old" && r.rm === row.rm);
        if (oldRow) {
          const oldVal = oldRow.data["Column 39"];
          const newVal = row.data["Column 39"];
          const partnerValues = [
            "Размещение делового партнера",
            "Размещение партнера",
            "Партнер",
          ];
          const isNowPartner = partnerValues.some(
            (v) => newVal && String(newVal).trim() === v,
          );
          const wasNotPartner = !partnerValues.some(
            (v) => oldVal && String(oldVal).trim() === v,
          );
          if (isNowPartner && wasNotPartner) {
            matches = true;
          }
        }
      }
      if (!matches) return false;
    }

    return true;
  });

  // Проверяем, активен ли фильтр по изменениям
  const isChangeFilterActive =
    filters.changeType.size > 0 || filters.toReserve || filters.toPartner;

  const resetFilters = () => {
    setFilters({
      status: new Set(["БЫЛО", "СТАЛО", "НОВАЯ", "УДАЛЕНА"]),
      address: new Set(),
      floor: new Set(),
      city: new Set(),
      quantity: new Set(),
      changeType: new Set(),
      toReserve: false,
      toPartner: false,
    });
  };

  const handleExport = () => {
    const { Workbook } = require("exceljs");

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Сравнение данных");

    // Стили
    const styles = {
      old: {
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF0F0" },
        },
      },
      new: {
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "F0FFF0" },
        },
      },
      added: {
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "F0F0FF" },
        },
      },
      deleted: {
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF0D0" },
        },
        font: { strike: true },
      },
      header: {
        font: { bold: true, color: { argb: "FFFFFFFF" } },
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF555555" },
        },
        alignment: { vertical: "middle", horizontal: "left" },
      },
    };

    // Заголовки
    const headers = displayColumns.map((col) => {
      if (col === "Status") return "Статус";
      if (col === "ChangeType") return "Тип изменения";
      return columnLabels[col] || col;
    });

    worksheet.addRow(headers);
    worksheet.getRow(1).height = 20;
    worksheet.getRow(1).eachCell((cell) => {
      Object.assign(cell, styles.header);
    });

    // Добавляем строки
    filteredRows.forEach((row) => {
      const rowData = displayColumns.map((col) => {
        if (col === "Status") {
          if (row.source === "new") return "НОВАЯ";
          if (row.source === "deleted") return "УДАЛЕНА";
          return { old: "БЫЛО", new: "СТАЛО", deleted: "УДАЛЕНА" }[row.type];
        }

        if (col === "ChangeType") {
          if (row.source === "new") return "Новая запись";
          if (row.source === "deleted") return "Запись удалена";
          if (row.type === "new") {
            const oldRow = allRows.find(
              (r) => r.type === "old" && r.rm === row.rm,
            );
            if (!oldRow) return "Изменено";
            const changedFields = requiredColumns.filter(
              (key) =>
                JSON.stringify(oldRow.data[key]) !==
                JSON.stringify(row.data[key]),
            );
            if (changedFields.length === 0) return "Без изменений";

            const sortedFields = [...changedFields].sort((a) =>
              a === "Column 39" ? -1 : 1,
            );
            return sortedFields
              .map((key) => {
                const label = columnLabels[key];
                const oldValue = oldRow.data[key];
                const newValue = row.data[key];
                const oldStr = oldValue == null ? "—" : String(oldValue);
                const newStr = newValue == null ? "—" : String(newValue);
                if (key === "Column 39") {
                  return `${label}: "${oldStr}" → "${newStr}"`;
                }
                return `${label}: ${newStr}`;
              })
              .join(", ");
          }
          return "—";
        }

        const value = row.data[col];
        if (value == null || value === "" || value === "null") return "—";
        return String(value);
      });

      const excelRow = worksheet.addRow(rowData);

      // Применяем стили
      if (row.source === "new") {
        excelRow.eachCell((cell) => Object.assign(cell, styles.added));
      } else if (row.source === "deleted") {
        excelRow.eachCell((cell) => Object.assign(cell, styles.deleted));
      } else if (row.type === "old") {
        excelRow.eachCell((cell) => Object.assign(cell, styles.old));
      } else if (row.type === "new") {
        excelRow.eachCell((cell) => Object.assign(cell, styles.new));
      }
    });

    // Автоподбор ширины
    worksheet.columns.forEach((column, i) => {
      const maxLength = Math.max(
        headers[i].length,
        ...filteredRows.map((row) => {
          const value = row.data[displayColumns[i]] || "";
          return String(value).length;
        }),
        10,
      );
      column.width = Math.min(maxLength + 2, 50);
    });

    // Экспорт
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `сравнение_РМ_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    });
  };

  return (
    <div
      className={`transition-all duration-300 ${fullscreen ? "fixed inset-0 z-50 overflow-auto bg-white p-2" : "mx-auto max-w-7xl bg-white p-4"}`}
    >
      {!fullscreen && (
        <h1 className="mb-4 text-2xl font-bold text-gray-800">
          Сравнение Excel-файлов
        </h1>
      )}

      <div className="mb-6 space-y-4 rounded-lg bg-gray-50 p-4">
        <div>
          <label className="mb-1 block font-medium text-gray-700">
            Первый файл (старый)
          </label>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={(e) => handleFileChange(e, setFirstFile)}
            className="w-full rounded border border-gray-300 px-3 py-2"
          />
        </div>
        <div>
          <label className="mb-1 block font-medium text-gray-700">
            Второй файл (новый)
          </label>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={(e) => handleFileChange(e, setSecondFile)}
            className="w-full rounded border border-gray-300 px-3 py-2"
          />
        </div>
        <button
          onClick={handleUpload}
          disabled={loading}
          className="rounded bg-blue-600 px-6 py-2 text-white transition hover:bg-blue-700 disabled:bg-gray-400"
        >
          {loading ? "Обработка..." : "Сравнить файлы"}
        </button>
      </div>

      {/* Фильтры */}
      {diffData.length > 0 && (
        <div className="mb-4 space-y-3 border-b border-t border-gray-200 bg-white px-2 py-4">
          <div className="flex items-start justify-between">
            <h3 className="text-sm font-semibold text-gray-800">Фильтры</h3>
            <button
              onClick={resetFilters}
              className="text-xs text-gray-500 underline hover:text-gray-700"
            >
              Сбросить всё
            </button>
            <button
              onClick={handleExport}
              className="text-xs font-medium text-blue-600 underline hover:text-blue-800"
            >
              📥 Скачать Excel
            </button>
            <button
              onClick={toggleFullscreen}
              className={`flex items-center gap-1 rounded px-2 py-1 text-xs font-medium ${
                fullscreen
                  ? "bg-orange-100 text-orange-800 hover:bg-orange-200"
                  : "bg-gray-100 text-gray-800 hover:bg-gray-200"
              } `}
              title={fullscreen ? "Вернуть в окно" : "Развернуть на весь экран"}
            >
              {fullscreen ? <>Выйти</> : <>Полный экран</>}
            </button>
          </div>

          {/* Статус */}
          <div>
            <label className="mb-1 block text-xs font-medium text-gray-700">
              Статус
            </label>
            <div className="flex flex-wrap gap-2">
              {(["БЫЛО", "СТАЛО", "НОВАЯ", "УДАЛЕНА"] as const).map(
                (status) => (
                  <label key={status} className="flex items-center text-xs">
                    <input
                      type="checkbox"
                      checked={filters.status.has(status)}
                      onChange={(e) => {
                        const newSet = new Set(filters.status);
                        e.target.checked
                          ? newSet.add(status)
                          : newSet.delete(status);
                        setFilters((prev) => ({ ...prev, status: newSet }));
                      }}
                      className="mr-1"
                    />
                    <span
                      className={`rounded-full px-2 py-0.5 text-xs font-medium text-white ${
                        status === "БЫЛО"
                          ? "bg-red-500"
                          : status === "СТАЛО"
                            ? "bg-green-500"
                            : status === "НОВАЯ"
                              ? "bg-blue-500"
                              : "bg-orange-500"
                      }`}
                    >
                      {status}
                    </span>
                  </label>
                ),
              )}
            </div>
          </div>

          {/* Тип изменения */}
          {uniqueChangeTypes.length > 0 && (
            <div>
              <label className="mb-1 block text-xs font-medium text-gray-700">
                Изменено поле
              </label>
              <div className="flex flex-wrap gap-2">
                {uniqueChangeTypes.map((type) => (
                  <label key={type} className="flex items-center text-xs">
                    <input
                      type="checkbox"
                      checked={filters.changeType.has(type)}
                      onChange={(e) => {
                        const newSet = new Set(filters.changeType);
                        e.target.checked
                          ? newSet.add(type)
                          : newSet.delete(type);
                        setFilters((prev) => ({ ...prev, changeType: newSet }));
                      }}
                      className="mr-1"
                    />
                    <span
                      className={`rounded px-1.5 py-0.5 text-xs font-medium ${
                        type === "Признак"
                          ? "bg-purple-100 text-purple-800"
                          : "bg-gray-100 text-gray-800"
                      }`}
                    >
                      {type}
                    </span>
                  </label>
                ))}
              </div>
            </div>
          )}

          {/* Специальные фильтры по Признаку */}
          <div>
            <label className="mb-1 block text-xs font-medium text-gray-700">
              Изменение Признака
            </label>
            <div className="flex flex-col gap-1 text-xs">
              <label className="flex items-center">
                <input
                  type="checkbox"
                  checked={filters.toReserve}
                  onChange={(e) =>
                    setFilters((prev) => ({
                      ...prev,
                      toReserve: e.target.checked,
                    }))
                  }
                  className="mr-1"
                />
                <span className="rounded bg-yellow-100 px-2 py-1 text-xs font-medium text-yellow-800">
                  Признак → Резерв
                </span>
              </label>
              <label className="flex items-center">
                <input
                  type="checkbox"
                  checked={filters.toPartner}
                  onChange={(e) =>
                    setFilters((prev) => ({
                      ...prev,
                      toPartner: e.target.checked,
                    }))
                  }
                  className="mr-1"
                />
                <span className="rounded bg-blue-100 px-2 py-1 text-xs font-medium text-blue-800">
                  Признак → Партнер
                </span>
              </label>
            </div>
          </div>

          {/* Адрес */}
          {uniqueAddresses.length > 0 && (
            <div>
              <label className="mb-1 block text-xs font-medium text-gray-700">
                Адрес
              </label>
              <div className="flex max-h-24 flex-wrap gap-2 overflow-y-auto">
                {uniqueAddresses.map((addr) => (
                  <label
                    key={addr}
                    className="flex items-center whitespace-nowrap text-xs"
                  >
                    <input
                      type="checkbox"
                      checked={filters.address.has(String(addr))}
                      onChange={(e) => {
                        const newSet = new Set(filters.address);
                        e.target.checked
                          ? newSet.add(String(addr))
                          : newSet.delete(String(addr));
                        setFilters((prev) => ({ ...prev, address: newSet }));
                      }}
                      className="mr-1"
                    />
                    {addr}
                  </label>
                ))}
              </div>
            </div>
          )}

          {/* Этаж */}
          {uniqueFloors.length > 0 && (
            <div>
              <label className="mb-1 block text-xs font-medium text-gray-700">
                Этаж
              </label>
              <div className="flex flex-wrap gap-2">
                {uniqueFloors.map((floor) => (
                  <label key={floor} className="flex items-center text-xs">
                    <input
                      type="checkbox"
                      checked={filters.floor.has(String(floor))}
                      onChange={(e) => {
                        const newSet = new Set(filters.floor);
                        e.target.checked
                          ? newSet.add(String(floor))
                          : newSet.delete(String(floor));
                        setFilters((prev) => ({ ...prev, floor: newSet }));
                      }}
                      className="mr-1"
                    />
                    {floor}
                  </label>
                ))}
              </div>
            </div>
          )}

          {/* Город */}
          {uniqueCities.length > 0 && (
            <div>
              <label className="mb-1 block text-xs font-medium text-gray-700">
                Город
              </label>
              <div className="flex flex-wrap gap-2">
                {uniqueCities.map((city) => (
                  <label key={city} className="flex items-center text-xs">
                    <input
                      type="checkbox"
                      checked={filters.city.has(city)}
                      onChange={(e) => {
                        const newSet = new Set(filters.city);
                        e.target.checked
                          ? newSet.add(city)
                          : newSet.delete(city);
                        setFilters((prev) => ({ ...prev, city: newSet }));
                      }}
                      className="mr-1"
                    />
                    {city}
                  </label>
                ))}
              </div>
            </div>
          )}

          {/* Кол-во */}
          {uniqueQuantities.length > 0 && (
            <div>
              <label className="mb-1 block text-xs font-medium text-gray-700">
                Занятость РМ
              </label>
              <div className="flex flex-wrap gap-2">
                {uniqueQuantities.map((qty) => (
                  <label key={qty} className="flex items-center text-xs">
                    <input
                      type="checkbox"
                      checked={filters.quantity.has(qty)}
                      onChange={(e) => {
                        const newSet = new Set(filters.quantity);
                        e.target.checked ? newSet.add(qty) : newSet.delete(qty);
                        setFilters((prev) => ({ ...prev, quantity: newSet }));
                      }}
                      className="mr-1"
                    />
                    {qty}
                  </label>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Таблица */}
      {diffData.length > 0 && (
        <div
          className={`overflow-x-auto ${fullscreen ? "max-h-[90vh] border-none" : "mt-2 max-h-screen rounded-lg border shadow-sm"}`}
        >
          {" "}
          <table className="min-w-full border-collapse text-sm">
            <thead>
              <tr className="bg-gray-100 text-xs uppercase text-gray-700">
                {displayColumns.map((col, index) => (
                  <th
                    key={col}
                    className={`sticky top-0 border border-gray-300 px-3 py-2 text-left text-center font-semibold ${index === 0 ? "left-0 z-50 bg-gray-100" : ""} ${index === 1 ? "z-45 left-[40px] bg-gray-100" : ""} ${index === 2 ? "left-[120px] z-40 bg-gray-100 shadow-md" : ""} ${index > 2 ? "z-10 bg-gray-100" : ""} `}
                    style={{
                      width:
                        index === 0
                          ? "40px"
                          : index === 1
                            ? "80px"
                            : index === 2
                              ? "180px"
                              : col === "ChangeType"
                                ? "280px"
                                : "120px",
                      minWidth: index === 0 ? "40px" : undefined,
                      backgroundColor: "#f3f4f6",
                      fontWeight: "600",
                      zIndex:
                        index === 0
                          ? "100"
                          : index === 1
                            ? "100"
                            : index === 2
                              ? "100"
                              : "30",
                    }}
                  >
                    {col === "Checked"
                      ? "✅"
                      : col === "Status"
                        ? "Статус"
                        : col === "ChangeType"
                          ? "Тип изменения"
                          : columnLabels[col]}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredRows.length === 0 ? (
                <tr>
                  <td
                    colSpan={displayColumns.length}
                    className="py-4 text-center text-gray-500"
                  >
                    Нет данных по выбранным фильтрам.
                  </td>
                </tr>
              ) : (
                filteredRows.map((row, idx) => {
                  const nextRow = filteredRows[idx + 1];
                  const isPartOfChange =
                    row.type === "old" &&
                    nextRow &&
                    nextRow.type === "new" &&
                    row.rm === nextRow.rm;

                  return (
                    <React.Fragment key={`row-${idx}`}>
                      <tr
                        className={
                          row.type === "old"
                            ? "hover:bg-red-25"
                            : row.type === "new"
                              ? "hover:bg-green-25"
                              : "hover:bg-red-25"
                        }
                      >
                        {displayColumns.map((col, cellIndex) => {
                          // === Ключ строки (для чекбокса) ===
                          const rmKey = row.rm;

                          // === Определение значения для статуса ===
                          let displayStatus = "—";
                          if (col === "Status") {
                            if (row.source === "new") {
                              displayStatus = "НОВАЯ";
                            } else if (row.source === "deleted") {
                              displayStatus = "УДАЛЕНА";
                            } else {
                              displayStatus = {
                                old: "БЫЛО",
                                new: "СТАЛО",
                                deleted: "УДАЛЕНА",
                              }[row.type];
                            }
                          }
                          const value =
                            col === "Status" ? displayStatus : row.data[col];

                          // === Логика для "Тип изменения" (только Признак) ===
                          let changeText = "—";
                          if (col === "ChangeType") {
                            if (row.source === "new") {
                              const newVal = row.data["Column 39"];
                              if (
                                newVal != null &&
                                String(newVal).trim() !== "" &&
                                String(newVal).trim() !== "null"
                              ) {
                                changeText = `Признак: → "${renderValue(newVal)}"`;
                              }
                            } else if (row.source === "deleted") {
                              const oldVal = row.data["Column 39"];
                              if (
                                oldVal != null &&
                                String(oldVal).trim() !== "" &&
                                String(oldVal).trim() !== "null"
                              ) {
                                changeText = `Признак: "${renderValue(oldVal)}" →`;
                              }
                            } else if (row.type === "new") {
                              // Это "СТАЛО" — ищем парную "БЫЛО"
                              const oldRow = allRows.find(
                                (r) => r.type === "old" && r.rm === row.rm,
                              );
                              if (oldRow) {
                                const oldVal = oldRow.data["Column 39"];
                                const newVal = row.data["Column 39"];
                                const oldStr = String(oldVal).trim();
                                const newStr = String(newVal).trim();

                                const hasOld =
                                  oldVal != null &&
                                  oldStr !== "" &&
                                  oldStr !== "null";
                                const hasNew =
                                  newVal != null &&
                                  newStr !== "" &&
                                  newStr !== "null";

                                if (hasOld && hasNew) {
                                  if (oldStr !== newStr) {
                                    changeText = `Признак: "${renderValue(oldVal)}" → "${renderValue(newVal)}"`;
                                  } else {
                                    changeText = `Признак: "${renderValue(newVal)}"`;
                                  }
                                } else if (hasOld) {
                                  changeText = `Признак: "${renderValue(oldVal)}" →`;
                                } else if (hasNew) {
                                  changeText = `Признак: → "${renderValue(newVal)}"`;
                                }
                              } else {
                                // Нет "БЫЛО" — возможно, ошибка, но можно оставить
                                const newVal = row.data["Column 39"];
                                if (
                                  newVal != null &&
                                  String(newVal).trim() !== ""
                                ) {
                                  changeText = `Признак: → "${renderValue(newVal)}"`;
                                }
                              }
                            }
                          }

                          // === Подсветка изменений (только если не в режиме фильтрации по изменениям) ===
                          const isChangeFilterActive =
                            filters.changeType.size > 0 ||
                            filters.toReserve ||
                            filters.toPartner;

                          const isChanged =
                            !isPartOfChange &&
                            row.type === "new" &&
                            col !== "Status" &&
                            col !== "ChangeType";
                          const oldValue = allRows[idx - 1]?.data[col];
                          const newValue = row.data[col];
                          const fieldChanged =
                            isChanged &&
                            row.type === "new" &&
                            oldValue !== undefined &&
                            JSON.stringify(oldValue) !==
                              JSON.stringify(newValue);

                          // === Жирное выделение при фильтрации по изменениям ===
                          const shouldHighlightBold =
                            isChangeFilterActive && row.type === "new";

                          let isTargetField = false;
                          if (
                            shouldHighlightBold &&
                            col !== "Status" &&
                            col !== "ChangeType" &&
                            col !== "Checked"
                          ) {
                            if (
                              row.source === "new" ||
                              row.source === "deleted"
                            ) {
                              // skip
                            } else {
                              const oldRow = allRows[idx - 1];
                              if (
                                oldRow &&
                                oldRow.type === "old" &&
                                oldRow.rm === row.rm
                              ) {
                                isTargetField =
                                  JSON.stringify(oldRow.data[col]) !==
                                  JSON.stringify(row.data[col]);
                              }
                            }
                          }

                          // === Определяем, отмечена ли строка ===
                          const isChecked = checkedRows.has(rmKey);

                          return (
                            <td
                              key={`cell-${col}`}
                              className={`border border-gray-300 px-2 py-1 text-sm ${
                                cellIndex === 0
                                  ? "sticky left-0 z-50 bg-white text-center"
                                  : cellIndex === 1
                                    ? "z-45 sticky left-[40px] bg-white font-bold"
                                    : cellIndex === 2
                                      ? "sticky left-[120px] z-40 bg-white"
                                      : "relative"
                              } ${cellIndex === 2 && fieldChanged && !isChangeFilterActive ? "bg-yellow-200" : ""} ${cellIndex > 2 && fieldChanged && !isChangeFilterActive ? "bg-yellow-200" : ""} ${isTargetField ? "font-bold text-blue-700" : ""} ${col === "ChangeType" ? "text-xs italic text-red-600" : ""} ${isChecked ? "bg-gray-50 text-gray-500 line-through" : ""} ${
                                row.type === "old" && !isChecked
                                  ? "bg-red-50 text-red-700"
                                  : row.type === "new" &&
                                      row.source !== "new" &&
                                      !isChecked
                                    ? "bg-green-50 text-green-700"
                                    : row.source === "new" && !isChecked
                                      ? "bg-blue-50 font-bold text-blue-800"
                                      : row.source === "deleted" && !isChecked
                                        ? "bg-red-50 text-red-500"
                                        : ""
                              } `}
                              style={{
                                width:
                                  cellIndex === 0
                                    ? "40px"
                                    : cellIndex === 1
                                      ? "80px"
                                      : cellIndex === 2
                                        ? "180px"
                                        : col === "ChangeType"
                                          ? "280px"
                                          : "120px",
                                minWidth: cellIndex === 0 ? "40px" : undefined,
                                top: 0,
                                height: "auto",
                                boxSizing: "border-box",
                                zIndex: cellIndex === 1 ? 25 : undefined,

                                background:
                                  cellIndex <= 2
                                    ? isChecked
                                      ? "#f8f8f8"
                                      : "white"
                                    : undefined,
                              }}
                            >
                              {col === "Checked" ? (
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => {
                                    setCheckedRows((prev) => {
                                      const newSet = new Set(prev);
                                      if (newSet.has(rmKey)) {
                                        newSet.delete(rmKey);
                                      } else {
                                        newSet.add(rmKey);
                                      }
                                      return newSet;
                                    });
                                  }}
                                  className="cursor-pointer"
                                  onClick={(e) => e.stopPropagation()}
                                />
                              ) : col === "ChangeType" ? (
                                changeText
                              ) : (
                                renderValue(value)
                              )}
                            </td>
                          );
                        })}
                      </tr>

                      {/* Отступ после пары БЫЛО+СТАЛО */}
                      {isPartOfChange && (
                        <tr>
                          <td
                            colSpan={displayColumns.length}
                            className="h-2"
                          ></td>
                        </tr>
                      )}

                      {/* Отступ после одиночной строки */}
                      {!isPartOfChange && (
                        <tr>
                          <td
                            colSpan={displayColumns.length}
                            className="h-2"
                          ></td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* Нет данных */}
      {!loading && diffData.length === 0 && (
        <p className="mt-6 text-gray-500">Загрузите два файла для сравнения.</p>
      )}
    </div>
  );
};

export default App;
