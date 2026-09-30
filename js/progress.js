// Shared UI for multi-document exports. A unit completes when it is added to the ZIP.
const ExportProgress = (() => {
    const panel = document.createElement('section');
    panel.className = 'export-progress';
    panel.hidden = true;
    panel.setAttribute('role', 'status');
    panel.setAttribute('aria-live', 'polite');
    panel.innerHTML = `
        <div class="export-progress__header">
            <span class="export-progress__icon" aria-hidden="true">↓</span>
            <div class="export-progress__text">
                <strong class="export-progress__title"></strong>
                <span class="export-progress__status"></span>
            </div>
            <button type="button" class="export-progress__close" aria-label="Закрыть уведомление" hidden>×</button>
        </div>
        <div class="export-progress__track" role="progressbar" aria-label="Готовность документов" aria-valuemin="0" aria-valuemax="1" aria-valuenow="0">
            <span class="export-progress__fill"></span>
        </div>`;
    document.body.appendChild(panel);

    const title = panel.querySelector('.export-progress__title');
    const status = panel.querySelector('.export-progress__status');
    const track = panel.querySelector('.export-progress__track');
    const fill = panel.querySelector('.export-progress__fill');
    const close = panel.querySelector('.export-progress__close');
    let total = 0;
    let completed = 0;
    let active = false;
    let hideTimer;
    let disabledButtons = [];

    function restoreButtons() {
        disabledButtons.forEach(([button, disabled]) => { button.disabled = disabled; });
        disabledButtons = [];
    }

    function start(expected, label) {
        if (active) return false;
        clearTimeout(hideTimer);
        total = Math.max(1, expected);
        completed = 0;
        active = true;
        panel.hidden = false;
        panel.classList.remove('export-progress--error');
        close.hidden = true;
        title.textContent = label;
        status.textContent = `Подготовка документов: 0 из ${total}`;
        track.setAttribute('aria-valuemax', String(total));
        track.setAttribute('aria-valuenow', '0');
        fill.style.width = '0%';
        disabledButtons = [...document.querySelectorAll('button[onclick^="generate"]')].map(button => [button, button.disabled]);
        disabledButtons.forEach(([button]) => { button.disabled = true; });
        return true;
    }

    function advance() {
        if (!active) return;
        completed = Math.min(total, completed + 1);
        status.textContent = `Готово ${completed} из ${total} документов`;
        track.setAttribute('aria-valuenow', String(completed));
        fill.style.width = `${completed / total * 100}%`;
    }

    function trackZip(zip) {
        if (!active) return () => {};
        const original = zip.file;
        zip.file = function (name, ...args) {
            const result = original.call(this, name, ...args);
            if (typeof name === 'string' && args.length) advance();
            return result;
        };
        return () => { zip.file = original; };
    }

    function packing() {
        if (active) status.textContent = `Упаковка архива: ${completed} из ${total} документов`;
    }

    function done() {
        if (!active) return;
        active = false;
        restoreButtons();
        title.textContent = 'Архив готов';
        status.textContent = `Сформировано ${completed} из ${total} документов`;
        close.hidden = false;
        hideTimer = setTimeout(() => { panel.hidden = true; }, 3500);
    }

    function fail(error) {
        if (!active) return;
        active = false;
        restoreButtons();
        panel.classList.add('export-progress--error');
        title.textContent = 'Не удалось сформировать архив';
        status.textContent = error?.message || String(error);
        close.hidden = false;
    }

    close.addEventListener('click', () => { panel.hidden = true; });
    return {start, advance, trackZip, packing, done, fail, isActive: () => active, completed: () => completed};
})();
