// Shared date conversion for Excel imports and Russian document dates.
function excelDateToISO(value, date1904 = false) {
    if (value == null || value === '') return '';
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return [value.getFullYear(), String(value.getMonth() + 1).padStart(2, '0'), String(value.getDate()).padStart(2, '0')].join('-');
    }
    if (typeof value === 'number' || /^\d+(?:\.\d+)?$/.test(String(value))) {
        const date = XLSX.SSF.parse_date_code(Number(value), {date1904});
        return date ? [date.y, String(date.m).padStart(2, '0'), String(date.d).padStart(2, '0')].join('-') : '';
    }
    const text = String(value).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
    const ru = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(text);
    return ru ? [ru[3], ru[2].padStart(2, '0'), ru[1].padStart(2, '0')].join('-') : '';
}

function formatDateRU(value) {
    if (!value) return '';
    const iso = /^\d{4}-\d{2}-\d{2}$/.test(value) ? value : excelDateToISO(value);
    return iso ? iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.' + iso.slice(0, 4) : '';
}

function readExcel(file, preferredSheet, requiredColumn) {
    return file.arrayBuffer().then(buffer => {
        const workbook = XLSX.read(buffer, {type: 'array'});
        const names = [preferredSheet, ...workbook.SheetNames.filter(name => name !== preferredSheet)];
        for (const name of names) {
            if (!workbook.Sheets[name]) continue;
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
            if (!rows.some(row => Object.hasOwn(row, requiredColumn))) continue;
            const byNumber = new Map(rows.map(row => [String(row[requiredColumn]).trim(), row]));
            return {rows, byNumber, date1904: !!workbook.Workbook?.WBProps?.date1904};
        }
        throw new Error('Не найден лист с колонкой «' + requiredColumn + '»');
    });
}
