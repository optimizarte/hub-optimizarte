// Añadir al final de main.js

// 1. Evento para el botón de inicio
document.getElementById('btn-empezar').addEventListener('click', () => {
    // Ocultar inicio, mostrar vela
    document.getElementById('pantalla-inicio').classList.replace('visible', 'oculta');
    document.getElementById('pantalla-vela').classList.replace('oculta', 'visible');
    
    // Llamamos a la función que enciende el micrófono (creada en el paso anterior)
    iniciarVela();
});

// 2. Actualizar la función apagarVela() para que haga la magia visual
function apagarVela() {
    isCandleLit = false;
    
    // Ocultar la llama y el resplandor de la vela
    document.getElementById('llama').style.display = 'none';
    document.getElementById('resplandor').style.display = 'none';
    
    // Darle 1 segundo de pausa para que vean que la han apagado, y mostrar el mensaje
    setTimeout(() => {
        document.getElementById('pantalla-vela').classList.replace('visible', 'oculta');
        document.getElementById('pantalla-mensaje').classList.replace('oculta', 'visible');
    }, 1200);
    
    // Apagar el micro
    if (audioContext && audioContext.state !== 'closed') {
        audioContext.close();
    }
}