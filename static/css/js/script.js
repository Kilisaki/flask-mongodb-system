// Модальное окно
const modal = document.getElementById("modal");
const modalBody = document.getElementById("modal-body");
const closeBtn = document.querySelector(".close");

function openModal(url, title) {
    fetch(url)
        .then(response => response.text())
        .then(html => {
            modalBody.innerHTML = html;
            modal.style.display = "block";
            document.querySelector(".modal-content h2").textContent = title;
        });
}

closeBtn.onclick = function() {
    modal.style.display = "none";
}

window.onclick = function(event) {
    if (event.target == modal) {
        modal.style.display = "none";
    }
}

// Подтверждение удаления
function confirmDelete() {
    return confirm("Вы уверены, что хотите удалить эту запись?");
}

// Инициализация Flatpickr
document.addEventListener('DOMContentLoaded', function() {
    flatpickr(".datepicker", {
        locale: "ru",
        dateFormat: "Y-m-d",
        allowInput: true
    });
});
