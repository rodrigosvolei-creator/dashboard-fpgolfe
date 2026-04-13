// Botao Imprimir/PDF - listener global
document.addEventListener('click', function(e) {
    if (e.target && e.target.id === 'btn-print-oxr') {
        setTimeout(function() { window.print(); }, 300);
    }
});
