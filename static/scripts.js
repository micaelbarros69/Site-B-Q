document.addEventListener('DOMContentLoaded', () => {
    const socket = io();
    const form = document.getElementById('upload-form');
    const iniciarAutomacaoBtn = document.getElementById('iniciar-automacao-btn');
    const tipoArquivoSelect = document.getElementById('tipo-arquivo');
    const folderInput = document.getElementById('folder');
    const messages = document.getElementById('messages');

    form.addEventListener('submit', (event) => {
        event.preventDefault();
        const files = document.querySelector('input[type="file"]').files;
        const folder = folderInput.value;
        const formData = new FormData();
        formData.append('folder', folder);
        for (let i = 0; i < files.length; i++) {
            formData.append('files[]', files[i]);
        }

        fetch('/upload', {
            method: 'POST',
            body: formData
        }).then(response => {
            if (response.ok) {
                appendMessage('Arquivos enviados com sucesso!', 'info');
            }
        });
    });

    iniciarAutomacaoBtn.addEventListener('click', () => {
        const tipo = tipoArquivoSelect.value;
        const folder = folderInput.value;
        socket.emit('processar_arquivos', { tipo: tipo, folder: folder });
    });

    socket.on('upload_complete', (data) => {
        const message = `Arquivos enviados para a pasta ${data.folder}: ${data.files.join(', ')}`;
        appendMessage(message, 'info');
    });

    socket.on('resultado', (data) => {
        const message = `Arquivo ${data.arquivo} do tipo ${data.tipo} na pasta ${data.folder} foi ${data.status}.`;
        appendMessage(message, data.status === 'anexado' ? 'success' : 'error');
    });

    socket.on('erro', (data) => {
        const message = `Erro: ${data.erro}`;
        appendMessage(message, 'error');
    });

    function appendMessage(message, type) {
        const messageDiv = document.createElement('div');
        messageDiv.classList.add('message', type);
        messageDiv.textContent = message;
        messages.appendChild(messageDiv);
        messages.scrollTop = messages.scrollHeight;
    }
});
