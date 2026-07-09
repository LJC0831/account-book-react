import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

const STORAGE_KEY = "basic-budget-transactions";
const FIXED_STORAGE_KEY = "basic-budget-fixed-transactions";

const categories = {
  expense: ["식비", "교통", "쇼핑", "주거", "통신", "문화", "의료", "보험", "저축", "카드사용", "기타"],
  income: ["급여", "용돈", "부수입", "이자", "기타"],
};

const chartColors = ["#0f766e", "#2563eb", "#f59e0b", "#db2777", "#7c3aed", "#475569", "#16a34a", "#dc2626"];

const formatCurrency = (value) =>
  new Intl.NumberFormat("ko-KR", {
    style: "currency",
    currency: "KRW",
    maximumFractionDigits: 0,
  }).format(value);

const getMonthKey = (dateString) => dateString.slice(0, 7);

const shiftMonthKey = (monthKey, offset) => {
  const [year, month] = monthKey.split("-").map(Number);
  const date = new Date(year, month - 1 + offset, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
};

const formatMonthLabel = (monthKey) => {
  const [year, month] = monthKey.split("-");
  return `${year}년 ${Number(month)}월`;
};

const parseAmount = (value) => Number(String(value).replaceAll(",", ""));

const formatAmountInput = (value) => {
  const numericText = String(value).replace(/\D/g, "");
  if (!numericText) return "";
  return new Intl.NumberFormat("ko-KR").format(Number(numericText));
};

const normalizeType = (value) => {
  const text = String(value || "").trim();
  if (text === "수입" || text.toLowerCase() === "income") return "income";
  return "expense";
};

const buildValidDateText = (year, month, day) => {
  const normalizedYear = String(year).padStart(4, "0");
  const normalizedMonth = String(month).padStart(2, "0");
  const normalizedDay = String(day).padStart(2, "0");
  const date = new Date(Number(normalizedYear), Number(normalizedMonth) - 1, Number(normalizedDay));

  if (
    date.getFullYear() === Number(normalizedYear) &&
    date.getMonth() === Number(normalizedMonth) - 1 &&
    date.getDate() === Number(normalizedDay)
  ) {
    return `${normalizedYear}-${normalizedMonth}-${normalizedDay}`;
  }

  return "";
};

const getLocalDateText = (date = new Date()) =>
  buildValidDateText(date.getFullYear(), date.getMonth() + 1, date.getDate());

const today = getLocalDateText();

const normalizeExcelDate = (value, XLSX) => {
  if (value === undefined || value === null || value === "") return "";

  if (value instanceof Date) {
    return buildValidDateText(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }

  const text = String(value).trim();
  const digitsOnly = text.replace(/\D/g, "");

  if (/^\d{8}$/.test(digitsOnly)) {
    const parsedText = buildValidDateText(
      digitsOnly.slice(0, 4),
      digitsOnly.slice(4, 6),
      digitsOnly.slice(6, 8),
    );
    if (parsedText) return parsedText;
  }

  if (typeof value === "number" && value > 20000 && value < 80000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    const parsedText = parsed ? buildValidDateText(parsed.y, parsed.m, parsed.d) : "";
    if (parsedText) return parsedText;
  }

  const matched = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
  if (matched) {
    const [, year, monthValue, dayValue] = matched;
    const parsedText = buildValidDateText(year, monthValue, dayValue);
    if (parsedText) return parsedText;
  }

  const monthFirstMatched = text.match(/^(\d{1,2})[-./](\d{1,2})[-./](\d{2}|\d{4})$/);
  if (monthFirstMatched) {
    const [, monthValue, dayValue, yearValue] = monthFirstMatched;
    const year = yearValue.length === 2 ? `20${yearValue}` : yearValue;
    const parsedText = buildValidDateText(year, monthValue, dayValue);
    if (parsedText) return parsedText;
  }

  const koreanMatched = text.match(/^(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일?$/);
  if (koreanMatched) {
    const [, year, monthValue, dayValue] = koreanMatched;
    const parsedText = buildValidDateText(year, monthValue, dayValue);
    if (parsedText) return parsedText;
  }

  return "";
};

const getHeaderIndex = (headers, names, fallbackIndex) => {
  const index = headers.findIndex((header) => names.includes(String(header).trim()));
  return index >= 0 ? index : fallbackIndex;
};

const isHeaderRow = (row) => {
  const values = row.map((cell) => String(cell).trim());
  return ["날짜", "거래일자", "일자"].some((name) => values.includes(name));
};

const getValidRows = (rows) =>
  rows.filter((row) => {
    const amount = parseAmount(row.amount);
    return row.date && row.category && Number.isFinite(amount) && amount > 0;
  });

const getValidCategory = (type, value) => {
  const category = String(value || "").trim();
  return categories[type].includes(category) ? category : categories[type][0];
};

const pageRoutes = {
  "/": "dashboard",
  "/entry": "entry",
  "/fixed": "fixed",
};

const routePaths = {
  dashboard: "/",
  entry: "/entry",
  fixed: "/fixed",
};

function getPageFromLocation() {
  return pageRoutes[window.location.pathname] || "dashboard";
}

const createEntryRow = () => ({
  id: crypto.randomUUID(),
  type: "expense",
  date: today,
  category: "식비",
  amount: "",
  memo: "",
});

const excelTemplateRows = [
  { 날짜: today, 유형: "지출", 카테고리: "식비", 금액: 12000, 메모: "점심 식사" },
  { 날짜: today, 유형: "수입", 카테고리: "급여", 금액: 2800000, 메모: "이번 달 급여" },
];

const createFixedRow = () => ({
  id: crypto.randomUUID(),
  type: "expense",
  day: "1",
  category: "주거",
  amount: "",
  memo: "",
});

const toEditableFixedRows = (items) =>
  items.map((item) => ({
    ...item,
    amount: formatAmountInput(item.amount),
  }));

const sampleTransactions = [
  {
    id: crypto.randomUUID(),
    type: "income",
    date: today,
    category: "급여",
    amount: 2800000,
    memo: "이번 달 급여",
  },
  {
    id: crypto.randomUUID(),
    type: "expense",
    date: today,
    category: "식비",
    amount: 365000,
    memo: "외식 및 장보기",
  },
  {
    id: crypto.randomUUID(),
    type: "expense",
    date: today,
    category: "주거",
    amount: 620000,
    memo: "월세",
  },
  {
    id: crypto.randomUUID(),
    type: "expense",
    date: today,
    category: "교통",
    amount: 82500,
    memo: "대중교통",
  },
  {
    id: crypto.randomUUID(),
    type: "expense",
    date: today,
    category: "문화",
    amount: 128000,
    memo: "영화 및 구독",
  },
];

function MonthlyTrendChart({ data }) {
  const maxVal = Math.max(...data.flatMap((d) => [d.income, d.expense]), 100000);

  return (
    <div className="trend-chart-container">
      <div className="trend-chart-header">
        <div className="chart-legend">
          <div className="legend-item">
            <span className="legend-dot income" />
            <span className="legend-label">수입</span>
          </div>
          <div className="legend-item">
            <span className="legend-dot expense" />
            <span className="legend-label">지출</span>
          </div>
        </div>
      </div>
      <div className="trend-chart-bars">
        {data.map((d) => {
          const incomeHeight = d.income > 0 ? `${Math.max((d.income / maxVal) * 100, 3)}%` : "0%";
          const expenseHeight = d.expense > 0 ? `${Math.max((d.expense / maxVal) * 100, 3)}%` : "0%";
          return (
            <div className="trend-chart-col" key={d.monthKey}>
              <div className="trend-bar-track">
                <div className="trend-bar-wrapper">
                  <div
                    className="trend-bar income"
                    style={{ height: incomeHeight }}
                  >
                    <span className="bar-tooltip">수입: {formatCurrency(d.income)}</span>
                  </div>
                  <div
                    className="trend-bar expense"
                    style={{ height: expenseHeight }}
                  >
                    <span className="bar-tooltip">지출: {formatCurrency(d.expense)}</span>
                  </div>
                </div>
              </div>
              <span className="trend-col-label">{d.label}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function loadTransactions() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) return sampleTransactions;

  try {
    const parsed = JSON.parse(stored);
    return Array.isArray(parsed) ? parsed : sampleTransactions;
  } catch {
    return sampleTransactions;
  }
}

function loadFixedItems() {
  const stored = localStorage.getItem(FIXED_STORAGE_KEY);
  if (!stored) return [];

  try {
    const parsed = JSON.parse(stored);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function getMonthlyFixedTransactions(fixedItems, monthKey) {
  const [year, month] = monthKey.split("-").map(Number);
  const lastDay = new Date(year, month, 0).getDate();

  return fixedItems.map((item) => {
    const day = Math.min(Number(item.day), lastDay);
    const date = `${monthKey}-${String(day).padStart(2, "0")}`;

    return {
      ...item,
      id: `fixed-${item.id}-${monthKey}`,
      date,
      amount: Number(item.amount),
      fixed: true,
    };
  });
}

function App() {
  const [page, setPage] = useState(getPageFromLocation);
  const [transactions, setTransactions] = useState(loadTransactions);
  const [fixedItems, setFixedItems] = useState(loadFixedItems);
  const [fixedRows, setFixedRows] = useState(() => {
    const storedItems = loadFixedItems();
    return storedItems.length > 0 ? toEditableFixedRows(storedItems) : [createFixedRow()];
  });
  const [selectedMonth, setSelectedMonth] = useState(today.slice(0, 7));
  const [filterType, setFilterType] = useState("all");
  const [filterCategory, setFilterCategory] = useState("all");
  const [searchText, setSearchText] = useState("");
  const [sortOrder, setSortOrder] = useState("latest");
  const [isTransactionsOpen, setIsTransactionsOpen] = useState(true);
  const [entryRows, setEntryRows] = useState(() => [createEntryRow(), createEntryRow(), createEntryRow()]);
  const [isUploadModalOpen, setIsUploadModalOpen] = useState(false);
  const [uploadMessage, setUploadMessage] = useState("");
  const [entrySaveMessage, setEntrySaveMessage] = useState("");
  const [lastSavedRows, setLastSavedRows] = useState([]);
  const [editingTransaction, setEditingTransaction] = useState(null);
  const [uploadPreview, setUploadPreview] = useState({
    rows: [],
    invalidRows: [],
  });

  useEffect(() => {
    const handlePopState = () => {
      setPage(getPageFromLocation());
    };

    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
  }, []);

  const navigateTo = (nextPage) => {
    const nextPath = routePaths[nextPage] || "/";
    if (window.location.pathname !== nextPath) {
      window.history.pushState({}, "", nextPath);
    }
    setPage(nextPage);
  };

  const saveTransactions = (nextTransactions) => {
    setTransactions(nextTransactions);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(nextTransactions));
  };

  const saveFixedItems = (nextFixedItems) => {
    setFixedItems(nextFixedItems);
    setFixedRows(nextFixedItems.length > 0 ? toEditableFixedRows(nextFixedItems) : [createFixedRow()]);
    localStorage.setItem(FIXED_STORAGE_KEY, JSON.stringify(nextFixedItems));
  };

  const fixedMonthTransactions = useMemo(
    () => getMonthlyFixedTransactions(fixedItems, selectedMonth),
    [fixedItems, selectedMonth],
  );
  const fixedExpenseTotal = fixedMonthTransactions
    .filter((transaction) => transaction.type === "expense")
    .reduce((sum, transaction) => sum + transaction.amount, 0);

  const previousMonth = shiftMonthKey(selectedMonth, -1);

  const previousMonthTransactions = useMemo(
    () => [
      ...getMonthlyFixedTransactions(fixedItems, previousMonth),
      ...transactions.filter((transaction) => getMonthKey(transaction.date) === previousMonth),
    ],
    [fixedItems, transactions, previousMonth],
  );

  const monthAllTransactions = useMemo(
    () => [
      ...fixedMonthTransactions,
      ...transactions.filter((transaction) => getMonthKey(transaction.date) === selectedMonth),
    ],
    [fixedMonthTransactions, transactions, selectedMonth],
  );

  const monthTransactions = useMemo(
    () => {
      const normalizedSearch = searchText.trim().toLowerCase();

      return monthAllTransactions
        .filter((transaction) => {
          if (filterType !== "all" && transaction.type !== filterType) return false;
          if (filterCategory !== "all" && transaction.category !== filterCategory) return false;
          if (!normalizedSearch) return true;

          const searchableText = [
            transaction.date,
            transaction.type === "income" ? "수입" : "지출",
            transaction.category,
            transaction.memo,
            String(transaction.amount),
            formatAmountInput(transaction.amount),
          ]
            .join(" ")
            .toLowerCase();

          return searchableText.includes(normalizedSearch);
        })
        .sort((a, b) => {
          if (sortOrder === "amount-desc") return b.amount - a.amount;
          if (sortOrder === "amount-asc") return a.amount - b.amount;
          return b.date.localeCompare(a.date);
        });
    },
    [monthAllTransactions, filterType, filterCategory, searchText, sortOrder],
  );

  const summary = useMemo(() => {
    return monthAllTransactions.reduce(
      (acc, transaction) => {
        acc[transaction.type] += transaction.amount;
        return acc;
      },
      { income: 0, expense: 0 },
    );
  }, [monthAllTransactions]);

  const monthlyTrendData = useMemo(() => {
    const trendMonths = [];
    for (let i = -5; i <= 0; i++) {
      trendMonths.push(shiftMonthKey(selectedMonth, i));
    }

    return trendMonths.map((mKey) => {
      const mTrans = transactions.filter((t) => getMonthKey(t.date) === mKey);
      const mFixed = getMonthlyFixedTransactions(fixedItems, mKey);

      const totalIncome = [...mTrans, ...mFixed]
        .filter((t) => t.type === "income")
        .reduce((sum, t) => sum + t.amount, 0);

      const totalExpense = [...mTrans, ...mFixed]
        .filter((t) => t.type === "expense")
        .reduce((sum, t) => sum + t.amount, 0);

      return {
        monthKey: mKey,
        label: `${Number(mKey.split("-")[1])}월`,
        income: totalIncome,
        expense: totalExpense,
      };
    });
  }, [transactions, fixedItems, selectedMonth]);

  const previousSummary = useMemo(() => {
    return previousMonthTransactions.reduce(
      (acc, transaction) => {
        acc[transaction.type] += transaction.amount;
        return acc;
      },
      { income: 0, expense: 0 },
    );
  }, [previousMonthTransactions]);

  const categoryTotals = useMemo(() => {
    return monthAllTransactions
      .filter((transaction) => transaction.type === "expense")
      .reduce((acc, transaction) => {
        acc[transaction.category] = (acc[transaction.category] || 0) + transaction.amount;
        return acc;
      }, {});
  }, [monthAllTransactions]);

  const previousCategoryTotals = useMemo(() => {
    return previousMonthTransactions
      .filter((transaction) => transaction.type === "expense")
      .reduce((acc, transaction) => {
        acc[transaction.category] = (acc[transaction.category] || 0) + transaction.amount;
        return acc;
      }, {});
  }, [previousMonthTransactions]);

  const categorySummary = useMemo(() => {
    return Object.entries(categoryTotals)
      .map(([category, amount], index) => ({ category, amount, color: chartColors[index % chartColors.length] }))
      .sort((a, b) => b.amount - a.amount);
  }, [categoryTotals]);

  const balance = summary.income - summary.expense;
  const variableExpenseTotal = Math.max(summary.expense - fixedExpenseTotal, 0);
  const fixedExpenseRate = summary.expense > 0 ? (fixedExpenseTotal / summary.expense) * 100 : 0;
  const spendingRate = summary.income > 0 ? Math.min((summary.expense / summary.income) * 100, 100) : 0;
  const maxCategoryAmount = Math.max(...categorySummary.map((item) => item.amount), 1);
  const topCategory = categorySummary[0];
  const mostIncreasedCategory = categories.expense
    .map((category) => ({
      category,
      amount: (categoryTotals[category] || 0) - (previousCategoryTotals[category] || 0),
    }))
    .filter((item) => item.amount > 0)
    .sort((a, b) => b.amount - a.amount)[0];

  const validEntryRows = getValidRows(entryRows);

  const draftTotal = validEntryRows.reduce((acc, row) => {
    const amount = parseAmount(row.amount);
    acc[row.type] += amount;
    return acc;
  }, { income: 0, expense: 0 });

  const selectedMonthManualTransactions = useMemo(
    () =>
      transactions
        .filter((transaction) => getMonthKey(transaction.date) === selectedMonth)
        .sort((a, b) => b.date.localeCompare(a.date)),
    [transactions, selectedMonth],
  );

  const validFixedRows = fixedRows.filter((row) => {
    const amount = parseAmount(row.amount);
    const day = Number(row.day);
    return row.category && Number.isFinite(amount) && amount > 0 && Number.isInteger(day) && day >= 1 && day <= 31;
  });

  const fixedTotal = fixedItems.reduce((acc, item) => {
    acc[item.type] += Number(item.amount);
    return acc;
  }, { income: 0, expense: 0 });

  const resetUploadState = () => {
    setUploadMessage("");
    setUploadPreview({ rows: [], invalidRows: [] });
  };

  const closeUploadModal = () => {
    setIsUploadModalOpen(false);
    resetUploadState();
  };

  const handleRowChange = (id, field, value) => {
    setEntrySaveMessage("");
    setEntryRows((currentRows) =>
      currentRows.map((row) => {
        if (row.id !== id) return row;

        if (field === "type") {
          return {
            ...row,
            type: value,
            category: categories[value][0],
          };
        }

        return {
          ...row,
          [field]: field === "amount" ? formatAmountInput(value) : value,
        };
      }),
    );
  };

  const handleAddRow = () => {
    setEntrySaveMessage("");
    setEntryRows((currentRows) => [...currentRows, createEntryRow()]);
  };

  const handleRemoveRow = (id) => {
    setEntrySaveMessage("");
    setEntryRows((currentRows) => {
      if (currentRows.length === 1) return currentRows;
      return currentRows.filter((row) => row.id !== id);
    });
  };

  const handleDownloadTemplate = async () => {
    const XLSX = await import("xlsx");
    const worksheet = XLSX.utils.json_to_sheet(excelTemplateRows);
    worksheet["!cols"] = [{ wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 14 }, { wch: 24 }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "사용내역");
    XLSX.writeFile(workbook, "가계부_사용내역_업로드_양식.xlsx");
  };

  const handleUploadExcel = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const XLSX = await import("xlsx");
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const sheetRows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: true,
        dateNF: "yyyy-mm-dd",
      });
      const firstRow = sheetRows[0] || [];
      const hasHeader = isHeaderRow(firstRow);
      const headers = hasHeader ? firstRow : ["날짜", "유형", "카테고리", "금액", "메모"];
      const dataRows = hasHeader ? sheetRows.slice(1) : sheetRows;

      const dateIndex = getHeaderIndex(headers, ["날짜", "거래일자", "일자"], 0);
      const typeIndex = getHeaderIndex(headers, ["유형", "수입/지출", "구분"], 1);
      const categoryIndex = getHeaderIndex(headers, ["카테고리", "분류"], 2);
      const amountIndex = getHeaderIndex(headers, ["금액"], 3);
      const incomeIndex = getHeaderIndex(headers, ["입금액", "입금"], -1);
      const expenseIndex = getHeaderIndex(headers, ["출금액", "출금"], -1);
      const memoIndex = getHeaderIndex(headers, ["메모", "적요", "내용"], 4);

      const parsedRows = dataRows.map((row, index) => {
          const incomeAmount = incomeIndex >= 0 ? parseAmount(row[incomeIndex]) : 0;
          const expenseAmount = expenseIndex >= 0 ? parseAmount(row[expenseIndex]) : 0;
          const type = incomeAmount > 0 ? "income" : normalizeType(row[typeIndex]);
          const amountValue = row[amountIndex] || (incomeAmount > 0 ? row[incomeIndex] : row[expenseIndex]);
          const date = normalizeExcelDate(row[dateIndex], XLSX);

          return {
            id: crypto.randomUUID(),
            type,
            date,
            category: getValidCategory(type, row[categoryIndex]),
            amount: formatAmountInput(amountValue),
            memo: String(row[memoIndex] || "").trim(),
            sourceRow: hasHeader ? index + 2 : index + 1,
          };
        });
      const nextRows = parsedRows.filter((row) => row.date && parseAmount(row.amount) > 0);
      const invalidRows = parsedRows.filter((row) => !row.date || parseAmount(row.amount) <= 0);

      setUploadPreview({ rows: nextRows, invalidRows });

      if (nextRows.length === 0) {
        const hasAnyDate = parsedRows.some((row) => row.date);
        const hasAnyAmount = parsedRows.some((row) => parseAmount(row.amount) > 0);
        const reason = !hasAnyDate
          ? "날짜를 읽지 못했습니다."
          : !hasAnyAmount
            ? "금액을 읽지 못했습니다."
            : "유효한 행이 없습니다.";
        setUploadMessage(`${reason} 날짜와 금액이 입력된 행을 확인해주세요.`);
        return;
      }

      setUploadMessage(
        `가져오기 가능 ${nextRows.length}건, 오류 ${invalidRows.length}건을 확인했습니다.`,
      );
    } catch {
      setUploadMessage("엑셀 파일을 읽지 못했습니다. .xlsx 파일인지 확인해주세요.");
    } finally {
      event.target.value = "";
    }
  };

  const applyUploadPreview = () => {
    if (uploadPreview.rows.length === 0) return;
    setEntryRows(uploadPreview.rows);
    setEntrySaveMessage("");
    setUploadMessage(`${uploadPreview.rows.length}건을 입력 표에 반영했습니다. 확인 후 저장해주세요.`);
    setIsUploadModalOpen(false);
  };

  const handleFixedRowChange = (id, field, value) => {
    setFixedRows((currentRows) =>
      currentRows.map((row) => {
        if (row.id !== id) return row;

        if (field === "type") {
          return {
            ...row,
            type: value,
            category: categories[value][0],
          };
        }

        return {
          ...row,
          [field]: field === "amount" ? formatAmountInput(value) : value,
        };
      }),
    );
  };

  const handleAddFixedRow = () => {
    setFixedRows((currentRows) => [...currentRows, createFixedRow()]);
  };

  const handleRemoveFixedRow = (id) => {
    setFixedRows((currentRows) => {
      if (currentRows.length === 1) return currentRows;
      return currentRows.filter((row) => row.id !== id);
    });
  };

  const handleSaveFixedRows = (event) => {
    event.preventDefault();

    const nextFixedItems = validFixedRows.map((row) => ({
      id: row.id,
      type: row.type,
      day: String(Number(row.day)),
      category: row.category,
      amount: parseAmount(row.amount),
      memo: row.memo.trim(),
    }));

    saveFixedItems(nextFixedItems);
    navigateTo("dashboard");
  };

  const handleSubmit = (event) => {
    event.preventDefault();

    const rowsToSave = getValidRows(entryRows);
    const nextTransactions = rowsToSave.map((row) => ({
      id: crypto.randomUUID(),
      type: row.type,
      date: row.date,
      category: row.category,
      amount: parseAmount(row.amount),
      memo: row.memo.trim(),
    }));

    if (nextTransactions.length === 0) {
      setEntrySaveMessage("저장할 수 있는 행이 없습니다. 날짜와 금액을 입력해주세요.");
      return;
    }

    saveTransactions([...nextTransactions, ...transactions]);
    setSelectedMonth(getMonthKey(nextTransactions[0].date));
    setLastSavedRows(nextTransactions);
    setEntryRows([createEntryRow(), createEntryRow(), createEntryRow()]);
    setEntrySaveMessage("");
    navigateTo("dashboard");
  };

  const handleDelete = (id) => {
    saveTransactions(transactions.filter((transaction) => transaction.id !== id));
  };

  const handleDeleteSelectedMonthTransactions = () => {
    if (selectedMonthManualTransactions.length === 0) return;

    const confirmed = window.confirm(`${formatMonthLabel(selectedMonth)} 저장 내역 ${selectedMonthManualTransactions.length}건을 모두 삭제할까요?`);
    if (!confirmed) return;

    saveTransactions(transactions.filter((transaction) => getMonthKey(transaction.date) !== selectedMonth));
    setLastSavedRows([]);
  };

  const openEditModal = (transaction) => {
    setEditingTransaction({
      ...transaction,
      amount: formatAmountInput(transaction.amount),
    });
  };

  const handleEditChange = (field, value) => {
    setEditingTransaction((current) => {
      if (!current) return current;

      if (field === "type") {
        return {
          ...current,
          type: value,
          category: categories[value][0],
        };
      }

      return {
        ...current,
        [field]: field === "amount" ? formatAmountInput(value) : value,
      };
    });
  };

  const handleSaveEdit = () => {
    if (!editingTransaction) return;

    const amount = parseAmount(editingTransaction.amount);
    if (!editingTransaction.date || !editingTransaction.category || !Number.isFinite(amount) || amount <= 0) return;

    const nextTransactions = transactions.map((transaction) =>
      transaction.id === editingTransaction.id
        ? {
            ...transaction,
            type: editingTransaction.type,
            date: editingTransaction.date,
            category: editingTransaction.category,
            amount,
            memo: editingTransaction.memo.trim(),
          }
        : transaction,
    );

    saveTransactions(nextTransactions);
    setSelectedMonth(getMonthKey(editingTransaction.date));
    setEditingTransaction(null);
  };

  return (
    <div className="app-container">
      <header className="global-header">
        <div className="header-inner">
          <div className="header-logo" onClick={() => navigateTo("dashboard")}>
            <span className="logo-icon">💸</span>
            <span className="logo-title">Ant가계부</span>
          </div>
          <nav className="global-nav">
            <button
              type="button"
              className={`nav-link ${page === "dashboard" ? "active" : ""}`}
              onClick={() => navigateTo("dashboard")}
            >
              대시보드 통계
            </button>
            <button
              type="button"
              className={`nav-link ${page === "entry" ? "active" : ""}`}
              onClick={() => navigateTo("entry")}
            >
              수입/지출 입력
            </button>
            <button
              type="button"
              className={`nav-link ${page === "fixed" ? "active" : ""}`}
              onClick={() => navigateTo("fixed")}
            >
              고정 내역 관리
            </button>
          </nav>
        </div>
      </header>

      <div className="page-wrapper">
        {page === "fixed" && (
          <main className="app-shell page-transition">
            <section className="page-header">
              <div>
                <p className="eyebrow dark">Fixed Transactions</p>
                <h1>고정 수입/지출 관리</h1>
                <p className="page-description">월급, 월세, 구독료처럼 매달 반복되는 항목을 등록하면 통계에 자동 반영됩니다.</p>
              </div>
            </section>

            <section className="entry-layout spreadsheet-layout">
              <form className="panel spreadsheet-panel elevated" onSubmit={handleSaveFixedRows}>
                <div className="sheet-actions">
                  <div>
                    <p>Monthly Template</p>
                    <h2>반복 항목 입력</h2>
                  </div>
                  <button className="ghost-button compact" type="button" onClick={handleAddFixedRow}>
                    행 추가
                  </button>
                </div>

                <div className="sheet-table-wrap">
                  <table className="sheet-table fixed-sheet-table">
                    <thead>
                      <tr>
                        <th>매월 일자</th>
                        <th>유형</th>
                        <th>카테고리</th>
                        <th>금액</th>
                        <th>메모</th>
                        <th>삭제</th>
                      </tr>
                    </thead>
                    <tbody>
                      {fixedRows.map((row) => (
                        <tr key={row.id}>
                          <td>
                            <input
                              type="number"
                              min="1"
                              max="31"
                              value={row.day}
                              onChange={(event) => handleFixedRowChange(row.id, "day", event.target.value)}
                            />
                          </td>
                          <td>
                            <select
                              value={row.type}
                              onChange={(event) => handleFixedRowChange(row.id, "type", event.target.value)}
                            >
                              <option value="expense">지출</option>
                              <option value="income">수입</option>
                            </select>
                          </td>
                          <td>
                            <select
                              value={row.category}
                              onChange={(event) => handleFixedRowChange(row.id, "category", event.target.value)}
                            >
                              {categories[row.type].map((category) => (
                                <option key={category} value={category}>
                                  {category}
                                </option>
                              ))}
                            </select>
                          </td>
                          <td>
                            <input
                              type="text"
                              inputMode="numeric"
                              min="1"
                              placeholder="0"
                              value={row.amount}
                              onChange={(event) => handleFixedRowChange(row.id, "amount", event.target.value)}
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              placeholder="예: 월세, 급여, 구독료"
                              value={row.memo}
                              onChange={(event) => handleFixedRowChange(row.id, "memo", event.target.value)}
                            />
                          </td>
                          <td>
                            <button className="icon-button" type="button" onClick={() => handleRemoveFixedRow(row.id)}>
                              삭제
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="sheet-footer">
                  <span>저장 가능 고정 항목 {validFixedRows.length}개</span>
                  <button className="primary-button" type="submit">
                    고정 항목 저장
                  </button>
                </div>
              </form>

              <aside className="entry-aside">
                <SummaryCard title="고정 수입" value={fixedTotal.income} tone="income" caption="매월 자동 반영될 수입" />
                <SummaryCard title="고정 지출" value={fixedTotal.expense} tone="expense" caption="매월 자동 반영될 지출" />
                <SummaryCard title="고정 순액" value={fixedTotal.income - fixedTotal.expense} tone={fixedTotal.income - fixedTotal.expense >= 0 ? "income" : "expense"} caption="고정 수입에서 고정 지출 차감" />
              </aside>
            </section>
          </main>
        )}

        {page === "entry" && (
          <main className="app-shell page-transition">
            <section className="page-header">
              <div>
                <p className="eyebrow dark">New Transaction</p>
                <h1>수입/지출 입력</h1>
                <p className="page-description">새 거래를 저장하면 선택한 월의 통계 화면에 바로 반영됩니다.</p>
              </div>
            </section>

            <section className="entry-layout spreadsheet-layout">
              <form className="panel spreadsheet-panel elevated" onSubmit={handleSubmit}>
                <div className="sheet-actions">
                  <div>
                    <p>Spreadsheet Input</p>
                    <h2>여러 내역을 한 번에 입력</h2>
                  </div>
                  <div className="sheet-button-group">
                    <button
                      className="upload-action-button"
                      type="button"
                      onClick={() => {
                        resetUploadState();
                        setIsUploadModalOpen(true);
                      }}
                    >
                      UPLOAD
                    </button>
                    <button className="ghost-button compact" type="button" onClick={handleAddRow}>
                      +
                    </button>
                  </div>
                </div>

                <div className="sheet-table-wrap">
                  <table className="sheet-table">
                    <thead>
                      <tr>
                        <th>날짜</th>
                        <th>유형</th>
                        <th>카테고리</th>
                        <th>금액</th>
                        <th>메모</th>
                        <th>삭제</th>
                      </tr>
                    </thead>
                    <tbody>
                      {entryRows.map((row) => (
                        <tr key={row.id}>
                          <td>
                            <input
                              type="date"
                              value={row.date}
                              onChange={(event) => handleRowChange(row.id, "date", event.target.value)}
                            />
                          </td>
                          <td>
                            <select
                              value={row.type}
                              onChange={(event) => handleRowChange(row.id, "type", event.target.value)}
                            >
                              <option value="expense">지출</option>
                              <option value="income">수입</option>
                            </select>
                          </td>
                          <td>
                            <select
                              value={row.category}
                              onChange={(event) => handleRowChange(row.id, "category", event.target.value)}
                            >
                              {categories[row.type].map((category) => (
                                <option key={category} value={category}>
                                  {category}
                                </option>
                              ))}
                            </select>
                          </td>
                          <td>
                            <input
                              type="text"
                              inputMode="numeric"
                              min="1"
                              placeholder="0"
                              value={row.amount}
                              onChange={(event) => handleRowChange(row.id, "amount", event.target.value)}
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              placeholder="메모"
                              value={row.memo}
                              onChange={(event) => handleRowChange(row.id, "memo", event.target.value)}
                            />
                          </td>
                          <td>
                            <button className="icon-button" type="button" onClick={() => handleRemoveRow(row.id)}>
                              삭제
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="sheet-footer">
                  <div className="sheet-footer-status">
                    <span>저장 가능 행 {validEntryRows.length}개</span>
                    {entrySaveMessage ? <strong>{entrySaveMessage}</strong> : null}
                  </div>
                  <button className="save-button" type="button" onClick={handleSubmit}>
                    저장
                  </button>
                </div>
              </form>

              <aside className="entry-aside">
                <SummaryCard title="이번 달 수입" value={summary.income} tone="income" caption="저장된 수입 합계" />
                <SummaryCard title="이번 달 지출" value={summary.expense} tone="expense" caption="저장된 지출 합계" />
                <SummaryCard title="입력 예정 수입" value={draftTotal.income} tone="income" caption="현재 표에 입력된 수입" />
                <SummaryCard title="입력 예정 지출" value={draftTotal.expense} tone="expense" caption="현재 표에 입력된 지출" />
                <section className="panel recent-save-panel">
                  <div className="panel-title saved-history-title">
                    <div>
                      <p>Saved History</p>
                      <h2>{formatMonthLabel(selectedMonth)} 저장 내역</h2>
                    </div>
                    <button
                      className="danger-button compact"
                      type="button"
                      disabled={selectedMonthManualTransactions.length === 0}
                      onClick={handleDeleteSelectedMonthTransactions}
                    >
                      전체삭제
                    </button>
                  </div>
                  <label className="saved-history-month">
                    조회 월
                    <input
                      type="month"
                      value={selectedMonth}
                      onChange={(event) => setSelectedMonth(event.target.value)}
                    />
                  </label>
                  {lastSavedRows.length > 0 ? (
                    <div className="recent-save-notice">
                      방금 저장한 내역 {lastSavedRows.length}건이 반영되었습니다.
                    </div>
                  ) : null}
                  <div className="recent-save-list saved-history-list">
                    {selectedMonthManualTransactions.length === 0 ? (
                      <p className="empty">저장된 내역이 없습니다.</p>
                    ) : (
                      selectedMonthManualTransactions.map((row) => (
                        <div className="recent-save-item saved-history-item" key={row.id}>
                          <div>
                            <span>{row.date}</span>
                            <strong>
                              {row.type === "income" ? "+" : "-"}
                              {formatCurrency(row.amount)}
                            </strong>
                            <small>
                              {row.category}
                              {row.memo ? ` / ${row.memo}` : ""}
                            </small>
                          </div>
                          <button type="button" onClick={() => handleDelete(row.id)}>
                            삭제
                          </button>
                        </div>
                      ))
                    )}
                  </div>
                </section>
              </aside>
            </section>

            {isUploadModalOpen ? (
              <div className="modal-backdrop" role="presentation">
                <section className="upload-modal" role="dialog" aria-modal="true" aria-labelledby="upload-title">
                  <div className="modal-title">
                    <div>
                      <p>Excel Import</p>
                      <h2 id="upload-title">사용내역 엑셀 업로드</h2>
                    </div>
                    <button className="modal-close-button" type="button" onClick={closeUploadModal}>
                      닫기
                    </button>
                  </div>

                  <div className="upload-guide">
                    <strong>업로드 양식</strong>
                    <span>날짜, 유형, 카테고리, 금액, 메모 컬럼을 사용합니다.</span>
                  </div>

                  <div className="modal-actions">
                    <button className="ghost-button" type="button" onClick={handleDownloadTemplate}>
                      엑셀 폼 다운로드
                    </button>
                    <label className="file-upload-button">
                      파일 선택
                      <input type="file" accept=".xlsx,.xls" onChange={handleUploadExcel} />
                    </label>
                  </div>

                  {uploadMessage ? <p className="upload-message">{uploadMessage}</p> : null}

                  {uploadPreview.rows.length > 0 || uploadPreview.invalidRows.length > 0 ? (
                    <div className="upload-preview">
                      <div className="preview-summary">
                        <span>반영 가능 {uploadPreview.rows.length}</span>
                        <span>오류 {uploadPreview.invalidRows.length}</span>
                      </div>

                      {uploadPreview.rows.length > 0 ? (
                        <div className="preview-table-wrap">
                          <table className="preview-table">
                            <thead>
                              <tr>
                                <th>날짜</th>
                                <th>유형</th>
                                <th>카테고리</th>
                                <th>금액</th>
                                <th>메모</th>
                              </tr>
                            </thead>
                            <tbody>
                              {uploadPreview.rows.slice(0, 6).map((row) => (
                                <tr key={row.id}>
                                  <td>{row.date}</td>
                                  <td>{row.type === "income" ? "수입" : "지출"}</td>
                                  <td>{row.category}</td>
                                  <td>{row.amount}</td>
                                  <td>{row.memo}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      ) : null}

                      <button className="save-button preview-apply-button" type="button" onClick={applyUploadPreview}>
                        표에 반영
                      </button>
                    </div>
                  ) : null}
                </section>
              </div>
            ) : null}
          </main>
        )}

        {page === "dashboard" && (
          <main className="app-shell page-transition">
            <section className="hero-panel dashboard-hero">
              <div className="hero-copy">
                <p className="eyebrow">Monthly Analytics</p>
                <h1>이번 달 지출이 어디에 집중됐는지 확인하세요</h1>
                <p className="hero-description">
                  카테고리별 지출 비중을 비교하고, 수입 대비 지출 흐름을 한 화면에서 점검할 수 있습니다.
                </p>
              </div>

              <div className="balance-card">
                <span>이번 달 잔액</span>
                <strong>{formatCurrency(balance)}</strong>
                <div className="balance-meta">
                  <small>수입 {formatCurrency(summary.income)}</small>
                  <small>지출 {formatCurrency(summary.expense)}</small>
                </div>
                <div className="rate-track" aria-label="수입 대비 지출 비율">
                  <span style={{ width: `${spendingRate}%` }} />
                </div>
              </div>
            </section>

            <section className="toolbar">
              <label className="month-picker">
                <span>조회 월</span>
                <div className="month-stepper">
                  <button type="button" aria-label="이전 달" onClick={() => setSelectedMonth((month) => shiftMonthKey(month, -1))}>
                    ‹
                  </button>
                  <input
                    type="month"
                    value={selectedMonth}
                    aria-label={formatMonthLabel(selectedMonth)}
                    onChange={(event) => setSelectedMonth(event.target.value)}
                  />
                  <button type="button" aria-label="다음 달" onClick={() => setSelectedMonth((month) => shiftMonthKey(month, 1))}>
                    ›
                  </button>
                </div>
              </label>
              <button className="primary-button compact" type="button" onClick={() => navigateTo("entry")}>
                내역 추가
              </button>
              <button className="fixed-action-button" type="button" onClick={() => navigateTo("fixed")}>
                <span>고정</span>
                고정 항목 관리
              </button>
            </section>

            <section className="summary-grid" aria-label="월별 요약">
              <SummaryCard title="수입" value={summary.income} tone="income" caption="이번 달 들어온 금액" />
              <SummaryCard title="지출" value={summary.expense} tone="expense" caption="이번 달 사용한 금액" />
              <SummaryCard title="잔액" value={balance} tone={balance >= 0 ? "income" : "expense"} caption="남은 현금 흐름" />
            </section>

            <section className="comparison-grid" aria-label="전월 대비 비교">
              <ComparisonCard
                title="수입 비교"
                current={summary.income}
                previous={previousSummary.income}
                caption={`${formatMonthLabel(previousMonth)} 대비`}
              />
              <ComparisonCard
                title="지출 비교"
                current={summary.expense}
                previous={previousSummary.expense}
                caption={`${formatMonthLabel(previousMonth)} 대비`}
                inverse
              />
            </section>

            <section className="panel trend-panel">
              <div className="panel-title">
                <div>
                  <p>Monthly Trend</p>
                  <h2>최근 6개월 소비 흐름</h2>
                </div>
              </div>
              <MonthlyTrendChart data={monthlyTrendData} />
            </section>

            <section className="panel report-panel">
              <div className="panel-title report-title">
                <div>
                  <p>Monthly Report</p>
                  <h2>{formatMonthLabel(selectedMonth)} 리포트</h2>
                </div>
              </div>
              <div className="report-grid">
                <ReportItem
                  label="최다 지출 카테고리"
                  value={topCategory ? topCategory.category : "-"}
                  detail={topCategory ? formatCurrency(topCategory.amount) : "지출 내역 없음"}
                />
                <ReportItem
                  label="전월 대비 증가"
                  value={mostIncreasedCategory ? mostIncreasedCategory.category : "-"}
                  detail={mostIncreasedCategory ? `+${formatCurrency(mostIncreasedCategory.amount)}` : "증가한 카테고리 없음"}
                />
                <ReportItem
                  label="고정 지출 비중"
                  value={`${fixedExpenseRate.toFixed(1)}%`}
                  detail={
                    fixedExpenseRate >= 50
                      ? `${formatCurrency(fixedExpenseTotal)} 고정비 주의`
                      : `${formatCurrency(fixedExpenseTotal)} / 변동 ${formatCurrency(variableExpenseTotal)}`
                  }
                  tone={fixedExpenseRate >= 50 ? "warning" : "normal"}
                />
              </div>
            </section>

            <section className="analytics-grid">
              <section className="panel main-chart-panel">
                <div className="panel-title">
                  <div>
                    <p>Spending Ranking</p>
                    <h2>카테고리별 지출 비교</h2>
                  </div>
                  {topCategory ? <strong className="top-badge">최다 지출: {topCategory.category}</strong> : null}
                </div>

                <div className="comparison-chart">
                  {categorySummary.length === 0 ? (
                    <p className="empty">이 달의 지출이 없습니다.</p>
                  ) : (
                    categorySummary.map((item) => {
                      const percent = summary.expense > 0 ? (item.amount / summary.expense) * 100 : 0;
                      const width = (item.amount / maxCategoryAmount) * 100;

                      return (
                        <div className="comparison-row" key={item.category}>
                          <div className="comparison-label">
                            <span className="legend-dot" style={{ background: item.color }} />
                            <strong>{item.category}</strong>
                            <small>{percent.toFixed(1)}%</small>
                          </div>
                          <div className="comparison-track">
                            <span style={{ width: `${width}%`, background: item.color }} />
                          </div>
                          <strong className="comparison-amount">{formatCurrency(item.amount)}</strong>
                        </div>
                      );
                    })
                  )}
                </div>
              </section>

              <aside className="panel insight-panel">
                <div className="panel-title">
                  <div>
                    <p>Quick Insight</p>
                    <h2>소비 요약</h2>
                  </div>
                </div>
                <div className="insight-list">
                  <InsightItem label="수입 대비 지출" value={`${spendingRate.toFixed(1)}%`} />
                  <InsightItem label="지출 카테고리 수" value={`${categorySummary.length}개`} />
                  <InsightItem label="가장 큰 지출" value={topCategory ? formatCurrency(topCategory.amount) : "-"} />
                </div>
              </aside>
            </section>

            <section className="panel transactions-panel">
              <div className="panel-title transactions-title">
                <div>
                  <p>History</p>
                  <h2>최근 거래 내역</h2>
                </div>
                <button className="collapse-button" type="button" onClick={() => setIsTransactionsOpen((open) => !open)}>
                  {isTransactionsOpen ? "접기" : `펼치기 (${monthTransactions.length})`}
                </button>
              </div>

              {isTransactionsOpen ? (
                <>
                <div className="transaction-filter-panel">
                  <div className="filter-buttons" aria-label="거래 유형 필터">
                    <button type="button" className={filterType === "all" ? "active" : ""} onClick={() => setFilterType("all")}>
                      전체
                    </button>
                    <button type="button" className={filterType === "expense" ? "active" : ""} onClick={() => setFilterType("expense")}>
                      지출
                    </button>
                    <button type="button" className={filterType === "income" ? "active" : ""} onClick={() => setFilterType("income")}>
                      수입
                    </button>
                  </div>
                  <div className="transaction-filter-grid">
                    <label>
                      검색
                      <input
                        type="search"
                        placeholder="메모, 카테고리, 금액"
                        value={searchText}
                        onChange={(event) => setSearchText(event.target.value)}
                      />
                    </label>
                    <label>
                      카테고리
                      <select value={filterCategory} onChange={(event) => setFilterCategory(event.target.value)}>
                        <option value="all">전체 카테고리</option>
                        {[...new Set([...categories.expense, ...categories.income])].map((category) => (
                          <option key={category} value={category}>
                            {category}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label>
                      정렬
                      <select value={sortOrder} onChange={(event) => setSortOrder(event.target.value)}>
                        <option value="latest">최신순</option>
                        <option value="amount-desc">금액 높은순</option>
                        <option value="amount-asc">금액 낮은순</option>
                      </select>
                    </label>
                  </div>
                </div>

                <div className="transaction-list">
                  {monthTransactions.length === 0 ? (
                    <p className="empty">조건에 맞는 내역이 없습니다.</p>
                  ) : (
                    monthTransactions.map((transaction) => (
                      <article className="transaction-item" key={transaction.id}>
                        <div className="transaction-main">
                          <span className={`type-dot ${transaction.type}`} />
                          <div>
                            <strong>{transaction.category}</strong>
                            <p>
                              {transaction.fixed ? "고정 항목 / " : ""}
                              {transaction.date}
                              {transaction.memo ? ` / ${transaction.memo}` : ""}
                            </p>
                          </div>
                        </div>
                        <div className="transaction-side">
                          <strong className={transaction.type}>
                            {transaction.type === "expense" ? "-" : "+"}
                            {formatCurrency(transaction.amount)}
                          </strong>
                          {transaction.fixed ? (
                            <button type="button" onClick={() => navigateTo("fixed")}>
                              관리
                            </button>
                          ) : (
                            <>
                              <button type="button" onClick={() => openEditModal(transaction)}>
                                수정
                              </button>
                              <button type="button" onClick={() => handleDelete(transaction.id)}>
                                삭제
                              </button>
                            </>
                          )}
                        </div>
                      </article>
                    ))
                  )}
                </div>
                </>
              ) : (
                <p className="collapsed-summary">
                  현재 조건에 맞는 거래 {monthTransactions.length}개가 접혀 있습니다.
                </p>
              )}
            </section>
          </main>
        )}
      </div>

      {editingTransaction ? (
        <div className="modal-backdrop" role="presentation">
          <section className="upload-modal edit-modal" role="dialog" aria-modal="true" aria-labelledby="edit-title">
            <div className="modal-title">
              <div>
                <p>Edit Transaction</p>
                <h2 id="edit-title">거래 내역 수정</h2>
              </div>
              <button className="modal-close-button" type="button" onClick={() => setEditingTransaction(null)}>
                닫기
              </button>
            </div>

            <div className="edit-form">
              <label>
                날짜
                <input
                  type="date"
                  value={editingTransaction.date}
                  onChange={(event) => handleEditChange("date", event.target.value)}
                />
              </label>
              <div className="field-grid">
                <label>
                  유형
                  <select
                    value={editingTransaction.type}
                    onChange={(event) => handleEditChange("type", event.target.value)}
                  >
                    <option value="expense">지출</option>
                    <option value="income">수입</option>
                  </select>
                </label>
                <label>
                  카테고리
                  <select
                    value={editingTransaction.category}
                    onChange={(event) => handleEditChange("category", event.target.value)}
                  >
                    {categories[editingTransaction.type].map((category) => (
                      <option key={category} value={category}>
                        {category}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              <label>
                금액
                <input
                  type="text"
                  inputMode="numeric"
                  value={editingTransaction.amount}
                  onChange={(event) => handleEditChange("amount", event.target.value)}
                />
              </label>
              <label>
                메모
                <input
                  type="text"
                  value={editingTransaction.memo}
                  onChange={(event) => handleEditChange("memo", event.target.value)}
                />
              </label>
            </div>

            <div className="modal-footer">
              <button className="ghost-button" type="button" onClick={() => setEditingTransaction(null)}>
                취소
              </button>
              <button className="save-button" type="button" onClick={handleSaveEdit}>
                저장
              </button>
            </div>
          </section>
        </div>
      ) : null}
    </div>
  );
}

function SummaryCard({ title, value, tone, caption }) {
  return (
    <article className={`summary-card ${tone}`}>
      <span>{title}</span>
      <strong>{formatCurrency(value)}</strong>
      <small>{caption}</small>
    </article>
  );
}

function ComparisonCard({ title, current, previous, caption, inverse = false }) {
  const diff = current - previous;
  const isPositive = diff > 0;
  const isNegative = diff < 0;
  const tone = inverse
    ? isPositive ? "bad" : isNegative ? "good" : "neutral"
    : isPositive ? "good" : isNegative ? "bad" : "neutral";

  return (
    <article className="comparison-card">
      <div>
        <span>{title}</span>
        <strong>{formatCurrency(current)}</strong>
      </div>
      <small>{caption}</small>
      <p className={tone}>
        {diff === 0 ? "전월과 동일" : `전월 대비 ${isPositive ? "+" : "-"}${formatCurrency(Math.abs(diff))}`}
      </p>
    </article>
  );
}

function ReportItem({ label, value, detail, tone = "normal" }) {
  return (
    <article className={`report-item ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
      <small>{detail}</small>
    </article>
  );
}

function InsightItem({ label, value }) {
  return (
    <div className="insight-item">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
