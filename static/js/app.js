document.querySelectorAll("form").forEach((form) => {
    form.addEventListener("submit", () => {
        const submitButton = form.querySelector("button[type='submit']");
        if (submitButton) {
            submitButton.disabled = true;
            submitButton.textContent = "Saving...";
        }
    });
});