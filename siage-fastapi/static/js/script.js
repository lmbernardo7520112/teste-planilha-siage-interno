document.addEventListener('DOMContentLoaded', () => {
    const uploadForm = document.getElementById('uploadForm');
    if (uploadForm) {
        uploadForm.addEventListener('submit', (event) => {
            const fileInput = document.getElementById('file');
            const file = fileInput.files[0];
            if (file && !file.name.endsWith('.json')) {
                event.preventDefault();
                alert('Please upload a valid JSON file.');
            }
        });
    }
});