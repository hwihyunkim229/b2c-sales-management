const modal = document.querySelector("#confirm-modal");
const modalTitle = document.querySelector("#confirm-modal-title");
const modalMessage = document.querySelector("#confirm-modal-message");
const cancelButton = document.querySelector("#confirm-modal-cancel");
const submitButton = document.querySelector("#confirm-modal-submit");
const sidebarToggle = document.querySelector(".sidebar-toggle");

if (sidebarToggle) {
    const storageKey = "b2c-sidebar-collapsed";
    const applySidebarState = (collapsed) => {
        document.body.classList.toggle("sidebar-collapsed", collapsed);
        document.documentElement.classList.remove("sidebar-collapsed-preload");
        sidebarToggle.setAttribute("aria-expanded", String(!collapsed));
        sidebarToggle.setAttribute(
            "aria-label",
            collapsed ? "사이드바 펼치기" : "사이드바 접기",
        );
    };

    applySidebarState(localStorage.getItem(storageKey) === "true");
    sidebarToggle.addEventListener("click", () => {
        const collapsed = !document.body.classList.contains("sidebar-collapsed");
        applySidebarState(collapsed);
        localStorage.setItem(storageKey, String(collapsed));
    });
}

const sizeTabs = Array.from(document.querySelectorAll("[data-size-tab]"));
const sizePanels = Array.from(document.querySelectorAll("[data-size-panel]"));
const scrollPreservingForms = document.querySelectorAll("form[data-preserve-scroll]");
const preservedScrollKey = `b2c-scroll:${window.location.pathname}`;

scrollPreservingForms.forEach((form) => {
    form.addEventListener("submit", () => {
        sessionStorage.setItem(preservedScrollKey, String(window.scrollY));
    });
});

const preservedScroll = sessionStorage.getItem(preservedScrollKey);
if (preservedScroll !== null) {
    sessionStorage.removeItem(preservedScrollKey);
    requestAnimationFrame(() => {
        requestAnimationFrame(() => {
            window.scrollTo(0, Number(preservedScroll));
        });
    });
}

sizeTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
        const selectedIndex = tab.dataset.sizeTab;
        sizeTabs.forEach((item) => {
            item.setAttribute(
                "aria-selected",
                String(item.dataset.sizeTab === selectedIndex),
            );
        });
        sizePanels.forEach((panel) => {
            panel.hidden = panel.dataset.sizePanel !== selectedIndex;
        });
    });
});

let pendingForm = null;

function openConfirmModal(form) {
    pendingForm = form;
    modalTitle.textContent =
        form.dataset.confirmTitle || "계속 진행할까요?";
    modalMessage.textContent =
        form.dataset.confirmMessage || "이 작업을 진행합니다.";
    submitButton.textContent =
        form.dataset.confirmLabel || "확인";

    modal.classList.add("is-open");
    modal.setAttribute("aria-hidden", "false");
    document.body.classList.add("modal-open");
    cancelButton.focus();
}

function closeConfirmModal() {
    modal.classList.remove("is-open");
    modal.setAttribute("aria-hidden", "true");
    document.body.classList.remove("modal-open");

    if (pendingForm) {
        const trigger = pendingForm.querySelector(
            "button[type='submit'], input[type='submit']",
        );

        if (trigger) {
            trigger.focus();
        }
    }

    pendingForm = null;
}

document.querySelectorAll("form[data-confirm-title]").forEach((form) => {
    form.addEventListener("submit", (event) => {
        event.preventDefault();
        openConfirmModal(form);
    });
});

cancelButton.addEventListener("click", closeConfirmModal);

submitButton.addEventListener("click", () => {
    if (!pendingForm) {
        return;
    }

    const formToSubmit = pendingForm;
    pendingForm = null;
    formToSubmit.submit();
});

modal.addEventListener("click", (event) => {
    if (event.target === modal) {
        closeConfirmModal();
    }
});

document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && modal.classList.contains("is-open")) {
        closeConfirmModal();
    }
});

function sortableValue(cell) {
    const text = cell.textContent.trim();
    const normalizedNumber = text
        .replaceAll(",", "")
        .replaceAll("원", "")
        .replaceAll("개", "")
        .replaceAll("행", "")
        .replaceAll("건", "")
        .trim();

    if (/^-?\d+(\.\d+)?$/.test(normalizedNumber)) {
        return {
            type: "number",
            value: Number(normalizedNumber),
        };
    }

    if (/^\d{4}-\d{2}$/.test(text)) {
        return {
            type: "date",
            value: text,
        };
    }

    return {
        type: "text",
        value: text.toLocaleLowerCase("ko-KR"),
    };
}

function compareCells(leftCell, rightCell, direction) {
    const left = sortableValue(leftCell);
    const right = sortableValue(rightCell);
    let result = 0;

    if (left.type === "number" && right.type === "number") {
        result = left.value - right.value;
    } else {
        result = String(left.value).localeCompare(
            String(right.value),
            "ko-KR",
            {
                numeric: true,
                sensitivity: "base",
            },
        );
    }

    return direction === "desc" ? -result : result;
}

document.querySelectorAll(".table-wrap table").forEach((table, tableIndex) => {
    if (table.dataset.serverSort === "true") {
        return;
    }

    const headers = Array.from(table.querySelectorAll("thead th"));
    const body = table.querySelector("tbody");

    if (!body || headers.length === 0) {
        return;
    }

    const toolbar = document.createElement("div");
    toolbar.className = "table-sort-toolbar";

    const toolbarLabel = document.createElement("span");
    toolbarLabel.textContent = "정렬";

    const columnSelect = document.createElement("select");
    columnSelect.setAttribute(
        "aria-label",
        `${tableIndex + 1}번째 테이블 정렬 기준`,
    );

    headers.forEach((header, columnIndex) => {
        const option = document.createElement("option");
        option.value = String(columnIndex);
        option.textContent = header.textContent.trim() || `컬럼 ${columnIndex + 1}`;
        columnSelect.append(option);
    });

    const directionSelect = document.createElement("select");
    directionSelect.setAttribute(
        "aria-label",
        `${tableIndex + 1}번째 테이블 정렬 방향`,
    );
    directionSelect.innerHTML = `
        <option value="asc">오름차순</option>
        <option value="desc">내림차순</option>
    `;

    const sortRows = () => {
        const columnIndex = Number(columnSelect.value);
        const direction = directionSelect.value;
        const rows = Array.from(body.querySelectorAll("tr")).filter(
            (row) => !row.querySelector("td.empty"),
        );

        rows.sort((leftRow, rightRow) => {
            return compareCells(
                leftRow.cells[columnIndex],
                rightRow.cells[columnIndex],
                direction,
            );
        });

        rows.forEach((row) => body.append(row));
    };

    columnSelect.addEventListener("change", sortRows);
    directionSelect.addEventListener("change", sortRows);

    toolbar.append(
        toolbarLabel,
        columnSelect,
        directionSelect,
    );
    table.parentElement.insertBefore(toolbar, table);
});
