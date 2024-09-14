import './App.css';
import * as XLSX from 'sheetjs-style';
import {useState} from "react";

function App() {
    const defaultValue = {
        name: '',
        entrance: '',
        exit: ''
    };
    const [waitresses, setWaitresses] = useState([Object.assign({}, defaultValue)]);

    function addWaitress() {
        setWaitresses([...waitresses, defaultValue])
    }

    function removeWaitress(i) {
        waitresses.splice(i, 1);
        setWaitresses([...waitresses]);
    }

    function changeName({target: elem}) {
        const {value} = elem;
        const [role, , id] = elem.name.split('-');

        if (role === 'waitress') {
            waitresses[id].name = value;
            setWaitresses([...waitresses]);
        }
    }

    function timeSelectChange({target: elem}) {
        const { value } = elem;
        if (value === '') return;

        let [hours, minutes] = value.split(':').map(a => parseInt(a));
        if (minutes < 10) minutes = `0${minutes}`;
        else if (minutes >= 60) {
            minutes = '00';
            hours++;
        }
        if (hours < 10) hours = `0${hours}`;

        const [role, input, id] = elem.name.split('-');

        if (role === 'waitress') {
            waitresses[id][input] = `${hours}:${minutes}`;
            setWaitresses([...waitresses]);
        }
    }

    function roundTime({target: elem}) {
        const { value } = elem;
        if (value === '') return;

        let [hours, minutes] = value.split(':').map(a => parseInt(a));
        const dif = minutes % 15;
        if (dif > 4)
            minutes += (15 - dif);
        else
            minutes -= dif;

        if (minutes < 10) minutes = `0${minutes}`;
        else if (minutes >= 60) {
            minutes = '00';
            hours++;
        }
        if (hours < 10) hours = `0${hours}`;

        const [role, input, id] = elem.name.split('-');

        if (role === 'waitress') {
            waitresses[id][input] = `${hours}:${minutes}`;
            setWaitresses([...waitresses]);
        }
    }

    function validateForm(form) {
        const DAY = 24;
        let formData = {};

        if (!form.reportValidity())
            return false;
        if (!form.date || !form.date.value)
            return form.reportValidity();
        if (!form['floor-tip'] || !form['floor-tip'].value)
            return form.reportValidity();
        if (!form['bar-tip'] || !form['bar-tip'].value)
            return form.reportValidity();

        formData.ftip = parseInt(form['floor-tip'].value);
        formData.btip = parseInt(form['bar-tip'].value);
        formData.date = form.date.value;

        formData.bartenders = [];
        formData.waitresses = [];

        formData.bh = 0;
        formData.wh = 0;
        let entrance, exit, diff;

        for (let i = 0, l = waitresses.length; i < l; i++) {
            entrance = new Date(`${formData.date}T${form[`waitress-entrance-${i}`].value}`);
            exit = new Date(`${formData.date}T${form[`waitress-exit-${i}`].value}`);
            diff = (exit - entrance) / 3600000;
            if (parseInt(form[`waitress-exit-${i}`].value.split(':')[0]) < 17) diff += DAY;
            formData.wh += diff;
            formData.waitresses.push([
                form[`waitress-name-${i}`].value,
                form[`waitress-entrance-${i}`].value,
                form[`waitress-exit-${i}`].value,
                diff
            ])
        }

        diff = 0;

        return formData;
    }

    function onSubmit(e) {
        e.preventDefault();
        const formData = validateForm(e.target);
    
        const totalTips = formData.ftip + formData.btip;
        formData.ftipPerHourBeforeSocial = totalTips / formData.wh;
        formData.ftipPerHour = (totalTips - formData.wh * 8) / formData.wh;
    
        const diffTo50 = formData.ftipPerHour >= 50 ? 0 : 50 - formData.ftipPerHour;
    
        const day = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"][new Date(formData.date).getDay()];
    
        let waitressesTopRow = [formData.ftip, formData.btip, totalTips, formData.wh, formData.ftipPerHour, formData.wh * 8];
        var data = [
            ["תאריך", "יום"],
            [formData.date, day],
            [],
            [],
            ["טיפים פלור", "טיפים בר", "סהכ טיפים", "סהכ שעות", "טיפ לשעה (אחרי הפרשות)", "סהכ הפרשות סוציאליות"],
            waitressesTopRow,
            [],
            ["שם", "שעת כניסה", "שעת יציאה", "סהכ שעות", "השלמה ל50", "סהכ טיפ", "הפרשות סוציאליות", "טיפ נטו"]
        ].concat(formData.waitresses.map(w => [
            ...w, 
            w[3] * diffTo50, 
            formData.ftipPerHourBeforeSocial * w[3], 
            w[3] * 8, 
            formData.ftipPerHour * w[3]
        ]));
    
        var workbook = XLSX.utils.book_new(),
            worksheet = XLSX.utils.aoa_to_sheet(data);
    
        // Apply styling to headers, second title row, and date cell
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell_address = { c: C, r: R };
                const cell_ref = XLSX.utils.encode_cell(cell_address);
                if (!worksheet[cell_ref]) continue;
    
                // Style for header rows
                if (R === 0 | R === 4 || R === 7) { // title rows
                    worksheet[cell_ref].s = {
                        font: { bold: true, color: { rgb: "FFFFFF" } },
                        fill: { fgColor: { rgb: "4F81BD" } },
                        alignment: { horizontal: "center" },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } }
                        }
                    };
                } else if (R === 0 && C === 1) { // Date cell
                    worksheet[cell_ref].s = {
                        font: { bold: true },
                        alignment: { horizontal: "center" }
                    };
                } else if (R > 3) { // Data rows
                    worksheet[cell_ref].s = {
                        alignment: { horizontal: "center" },
                        border: {
                            left: { style: "thin", color: { rgb: "D9D9D9" } },
                            right: { style: "thin", color: { rgb: "D9D9D9" } }
                        }
                    };
                } else {
                    worksheet[cell_ref].s = {
                        alignment: { horizontal: "center" }
                    };
                }
            }
        }
    
        // Set column widths, making column E wider
        worksheet['!cols'] = [
            { wch: 15 }, // תאריך
            { wch: 10 }, // Value column
            { wch: 10 }, // יום
            { wch: 10 }, // Value column
            { wch: 25 }, // Wider column for 'טיפ לשעה (אחרי הפרשות)'
            { wch: 20 },  // Other data columns
            { wch: 18 }  // Other data columns
        ];
    
        workbook.SheetNames.push("First");
        workbook.Sheets["First"] = worksheet;
    
        // Convert to binary string
        var xlsbin = XLSX.write(workbook, {
            bookType: "xlsx",
            type: "binary",
            cellStyles: true // Enable cell styles
        });
    
        // Convert to blob object
        var buffer = new ArrayBuffer(xlsbin.length),
            array = new Uint8Array(buffer);
        for (var i = 0; i < xlsbin.length; i++) {
            array[i] = xlsbin.charCodeAt(i) & 0xFF;
        }
        var xlsblob = new Blob([buffer], { type: "application/octet-stream" });
    
        var url = window.URL.createObjectURL(xlsblob),
            anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = `טיפים-${formData.date}.xlsx`;
        anchor.click();
        window.URL.revokeObjectURL(url);
    
        e.target.reset();
    }
    
    

    return (
        <div className="App">
            <header className="text-center my-5">
                <img src={process.env.PUBLIC_URL + '/logo.png'} alt="Logo" className="img-fluid mb-3 logo" />
                <h2>EZ Tip App</h2>
            </header>
            <div className="form-wrapper"> {/* Add this wrapper to center the form */}
                <form onSubmit={onSubmit} className="form-container p-4 shadow rounded">
                    <h3 className="section-title">תאריך</h3>
                    <div className="form-group mb-4">
                        <input type="date" name="date" id="date" className="form-control mb-3" required />
                    </div>
    
                    <section id="tips">
                        <h3 className="section-title">סה״כ טיפים פלור ובר</h3>
                        <div className="d-flex justify-content-between mb-4 gap-3"> {/* Add gap-3 class for spacing */}
                            <input placeholder="טיפים פלור" type="number" id="floor-tip" className="form-control" required />
                            <input placeholder="טיפים בר" type="number" id="bar-tip" className="form-control" required />
                        </div>
                    </section>
    
                    <section id="waitresses">
                        <h3 className="section-title">מלצריות וברמנים</h3>
                        <div className="list mb-4">
                            {waitresses.map((w, index) => (
                                <div key={index} className="waitress-item d-flex align-items-center">
                                    <input type="text" name={`waitress-name-${index}`} id={`waitress-name-${index}`}
                                        onChange={changeName}
                                        className="form-control me-2"
                                        placeholder="שם העובד" required />
                                    <input type="time" name={`waitress-entrance-${index}`} id={`waitress-entrance-${index}`}
                                        value={w.entrance} onChange={timeSelectChange} onBlur={roundTime}
                                        className="form-control me-2" required />
                                    <input type="time" name={`waitress-exit-${index}`} id={`waitress-exit-${index}`}
                                        value={w.exit} onChange={timeSelectChange} onBlur={roundTime}
                                        className="form-control me-2" required />
                                    <button type="button" className="btn btn-danger btn-sm" onClick={() => removeWaitress(index)}>
                                        <i className="fas fa-trash-alt"></i>
                                    </button>
                                </div>
                            ))}
                        </div>
                        <button type="button" className="btn btn-primary mb-4" onClick={addWaitress}>הוסף עובד</button>
                    </section>
    
                    <button type="submit" className="btn btn-success btn-lg w-100">סיימתי, חשב!</button>
                </form>
            </div>
            <footer className="text-center mt-4">
                <p className="text-muted">Made with love in Tel Aviv ❤️ by Tal Hazi</p>
            </footer>
        </div>
    );
    
    
}

export default App;
