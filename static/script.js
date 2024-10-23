function enviarSolicitud(nombreUniversidad) {
    fetch('/procesar', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ universidad: nombreUniversidad })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`Network response was not ok. Status: ${response.status} - ${response.statusText}`);
        }
        return response.json();
    })
    .then(data => {
        console.log('Success:', data);
        if (data.status === 'success') {
            if (Array.isArray(data.resultados) && data.resultados.length > 0) {
                data.resultados.forEach(universidad => {
                    console.log(`Resultado para ${universidad.nombre}: ${universidad.resultado}`);
                    // Aquí puedes realizar cualquier acción adicional con los resultados
                    alert(`Resultado para ${universidad.nombre}: ${universidad.resultado}`);
                });
            } else {
                console.error('No se encontraron resultados para mostrar.');
            }
        } else {
            console.error('Error en el servidor:', data.message);
        }
    })
    .catch(error => {
        console.error('There has been a problem with your fetch operation:', error);
        alert(`Error: ${error.message}`);
    });
}
