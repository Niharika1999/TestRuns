document.getElementById('assessmentForm').addEventListener('submit', function(event) {
    event.preventDefault();
    
    const formData = new FormData(event.target);
    const formObject = Object.fromEntries(formData.entries());
    formObject.services = formData.getAll('services'); // Get all selected services

    fetch('http://localhost:3000/update-excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(formObject),
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            alert('Form submitted and Excel updated successfully!');
            event.target.reset(); // Reset form fields
        } else {
            alert('Error updating Excel file.');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('Error updating Excel file.');
    });
});
