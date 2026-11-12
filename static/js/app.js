document.querySelectorAll("form").forEach((form) => {
    form.addEventListener("click", (event) => {
        const removeButton = event.target.closest("[data-remove-row]");
        if (removeButton) {
            const list = removeButton.closest("[data-repeatable-list]");
            if (list && list.children.length > 1) {
                removeButton.closest("fieldset, .skill-row").remove();
            }
        }

        const addButton = event.target.closest("[data-add-row]");
        if (addButton) {
            const list = form.querySelector("[data-repeatable-list]");
            const rowTemplate = form.querySelector(`[data-row-template="${addButton.dataset.addRow}"]`);
            if (list && rowTemplate) {
                const row = rowTemplate.content.cloneNode(true);
                list.append(row);
                const firstInput = list.lastElementChild.querySelector("input, textarea");
                if (firstInput) firstInput.focus();
            }
        }
    });

    form.addEventListener("submit", () => {
        const submitButton = form.querySelector("button[type='submit']");
        if (submitButton) {
            submitButton.disabled = true;
            submitButton.textContent = "Saving...";
        }
    });
});