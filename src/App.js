import './App.css';
import * as XLSX from 'sheetjs-style';
import { useReducer, useCallback, useState } from 'react';

// State reducer for managing waitress actions
function waitressReducer(state, action) {
  switch (action.type) {
    case 'ADD_WAITRESS':
      return [...state, { name: '', entrance: '', exit: '' }];
    case 'REMOVE_WAITRESS':
      return state.filter((_, index) => index !== action.index);
    case 'UPDATE_WAITRESS':
      return state.map((waitress, index) =>
        index === action.index ? { ...waitress, [action.field]: action.value } : waitress
      );
    default:
      return state;
  }
}

function App() {
  const [waitresses, dispatch] = useReducer(waitressReducer, [{ name: '', entrance: '', exit: '' }]);
  const [isDarkTheme, setIsDarkTheme] = useState(false);

  const toggleTheme = () => {
    setIsDarkTheme((prev) => !prev);
  };

  const addWaitress = useCallback(() => dispatch({ type: 'ADD_WAITRESS' }), []);
  const removeWaitress = useCallback((index) => dispatch({ type: 'REMOVE_WAITRESS', index }), []);
  const updateWaitress = useCallback((index, field, value) => {
    dispatch({ type: 'UPDATE_WAITRESS', index, field, value });
  }, []);

  const validateForm = useCallback(
    (form) => {
      const DAY = 24;
      let formData = {};

      if (!form.reportValidity()) return false;
      if (!form.date || !form.date.value) return form.reportValidity();
      if (!form['floor-tip'] || !form['floor-tip'].value) return form.reportValidity();
      if (!form['bar-tip'] || !form['bar-tip'].value) return form.reportValidity();

      formData.ftip = parseInt(form['floor-tip'].value);
      formData.btip = parseInt(form['bar-tip'].value);
      formData.date = form.date.value;

      formData.waitresses = [];
      formData.wh = 0;
      let entrance, exit, diff;

      waitresses.forEach((_, i) => {
        entrance = new Date(`${formData.date}T${form[`waitress-entrance-${i}`].value}`);
        exit = new Date(`${formData.date}T${form[`waitress-exit-${i}`].value}`);
        diff = (exit - entrance) / 3600000;
        if (parseInt(form[`waitress-exit-${i}`].value.split(':')[0]) < 17) diff += DAY;
        formData.wh += diff;
        formData.waitresses.push([
          form[`waitress-name-${i}`].value,
          form[`waitress-entrance-${i}`].value,
          form[`waitress-exit-${i}`].value,
          diff,
        ]);
      });

      return formData;
    },
    [waitresses]
  );

  const createExcelSheet = useCallback((formData) => {
    const totalTips = formData.ftip + formData.btip;
    formData.ftipPerHourBeforeSocial = totalTips / formData.wh;
    formData.ftipPerHour = (totalTips - formData.wh * 8) / formData.wh;

    const diffTo50 = formData.ftipPerHour >= 50 ? 0 : 50 - formData.ftipPerHour;

    const day = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'][new Date(formData.date).getDay()];

    let waitressesTopRow = [
      formData.ftip,
      formData.btip,
      totalTips,
      formData.wh,
      formData.ftipPerHour.toFixed(2),
      formData.wh * 8,
    ];
    const data = [
      ['תאריך', 'יום'],
      [formData.date, day],
      [],
      [],
      ['טיפים פלור', 'טיפים בר', 'סהכ טיפים', 'סהכ שעות', 'טיפ לשעה (אחרי הפרשות)', 'סהכ הפרשות סוציאליות'],
      waitressesTopRow,
      [],
      ['שם', 'שעת כניסה', 'שעת יציאה', 'סהכ שעות', 'השלמה ל50', 'סהכ טיפ', 'הפרשות סוציאליות', 'טיפ נטו'],
    ].concat(
      formData.waitresses.map((w) => [
        ...w,
        w[3] * diffTo50,
        (formData.ftipPerHourBeforeSocial * w[3]).toFixed(2),
        w[3] * 8,
        (formData.ftipPerHour * w[3]).toFixed(2),
      ])
    );

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // Apply styling to headers and data cells
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = { c: C, r: R };
        const cellRef = XLSX.utils.encode_cell(cellAddress);
        if (!worksheet[cellRef]) continue;

        // Style headers
        if ([0, 4, 7].includes(R)) {
          worksheet[cellRef].s = {
            font: { bold: true, color: { rgb: 'FFFFFF' } },
            fill: { fgColor: { rgb: '4F81BD' } },
            alignment: { horizontal: 'center' },
            border: {
              top: { style: 'thin', color: { rgb: '000000' } },
              bottom: { style: 'thin', color: { rgb: '000000' } },
            },
          };
        } else {
          worksheet[cellRef].s = {
            alignment: { horizontal: 'center' },
            border: {
              left: { style: 'thin', color: { rgb: 'D9D9D9' } },
              right: { style: 'thin', color: { rgb: 'D9D9D9' } },
            },
          };
        }
      }
    }

    // Set column widths
    worksheet['!cols'] = [
      { wch: 15 }, // תאריך
      { wch: 10 }, // יום
      { wch: 20 }, // Wider column
      { wch: 20 }, // Other data columns
      { wch: 25 }, // Other data columns
      { wch: 20 }, // Other data columns
      { wch: 20 }, // Other data columns
    ];

    workbook.SheetNames.push('Tips');
    workbook.Sheets['Tips'] = worksheet;

    const xlsbin = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'binary',
      cellStyles: true,
    });

    const buffer = new ArrayBuffer(xlsbin.length);
    const array = new Uint8Array(buffer);
    for (let i = 0; i < xlsbin.length; i++) {
      array[i] = xlsbin.charCodeAt(i) & 0xff;
    }
    const xlsblob = new Blob([buffer], { type: 'application/octet-stream' });

    const url = window.URL.createObjectURL(xlsblob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `טיפים-${formData.date}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  }, []);

  const onSubmit = useCallback(
    (e) => {
      e.preventDefault();
      const formData = validateForm(e.target);
      if (formData) {
        createExcelSheet(formData);
        e.target.reset();
      }
    },
    [validateForm, createExcelSheet]
  );

  return (
    <div className={`App ${isDarkTheme ? 'dark-theme' : ''}`}>
      <header className="text-center my-5">
        <img src={`${process.env.PUBLIC_URL}/logo.png`} alt="Logo" className="img-fluid mb-3 logo" />
        <h2>EZ Tip App</h2>
        <div className="theme-toggle">
          <label className="switch">
            <input type="checkbox" checked={isDarkTheme} onChange={toggleTheme} />
            <span className="slider"></span>
            <span className="toggle-label">{isDarkTheme ? 'Dark Dark' : 'Light Light'}</span>
          </label>
        </div>
      </header>
      <div className="form-wrapper">
        <form onSubmit={onSubmit} className="form-container p-4 shadow rounded">
          <h3 className="section-title">תאריך</h3>
          <div className="form-group mb-4">
            <input type="date" name="date" id="date" className="form-control mb-3" required />
          </div>

          <section id="tips">
            <h3 className="section-title">סה״כ טיפים פלור ובר</h3>
            <div className="d-flex justify-content-between mb-4 gap-3">
              <input placeholder="טיפים פלור" type="number" id="floor-tip" className="form-control" required />
              <input placeholder="טיפים בר" type="number" id="bar-tip" className="form-control" required />
            </div>
          </section>

          <section id="waitresses">
            <h3 className="section-title">מלצריות וברמנים</h3>
            <div className="list mb-4">
              {waitresses.map((w, index) => (
                <div key={index} className="waitress-item d-flex align-items-center">
                  <input
                    type="text"
                    name={`waitress-name-${index}`}
                    value={w.name}
                    onChange={(e) => updateWaitress(index, 'name', e.target.value)}
                    className="form-control me-2"
                    placeholder="שם העובד"
                    required
                  />
                  <input
                    type="time"
                    name={`waitress-entrance-${index}`}
                    value={w.entrance}
                    onChange={(e) => updateWaitress(index, 'entrance', e.target.value)}
                    className="form-control me-2"
                    required
                  />
                  <input
                    type="time"
                    name={`waitress-exit-${index}`}
                    value={w.exit}
                    onChange={(e) => updateWaitress(index, 'exit', e.target.value)}
                    className="form-control me-2"
                    required
                  />
                  <button type="button" className="btn btn-danger btn-sm" onClick={() => removeWaitress(index)}>
                    <i className="fas fa-times"></i>
                  </button>
                </div>
              ))}
            </div>
            <button type="button" className="btn btn-primary mb-4" onClick={addWaitress}>
              <i className="fas fa-plus"></i> הוסף עובד
            </button>
          </section>

          <button type="submit" className="btn btn-success btn-lg w-100">
            סיימתי, חשב!
          </button>
        </form>
      </div>
      <footer className="text-center mt-4">
        <p className="footer-text">Made with love in Tel Aviv ❤️ by Tal Hazi</p>
      </footer>
    </div>
  );
}

export default App;
