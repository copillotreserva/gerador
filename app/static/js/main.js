let certificados = [];

function formatarData(input) {
    let v = input.value.replace(/\D/g, '').slice(0, 8);
    if (v.length >= 5) {
        input.value = `${v.slice(0, 2)}/${v.slice(2, 4)}/${v.slice(4)}`;
    } else if (v.length >= 3) {
        input.value = `${v.slice(0, 2)}/${v.slice(2)}`;
    } else {
        input.value = v;
    }
}

function formatarCertificado(input) {
    let v = input.value.replace(/\D/g, '').slice(0, 8);
    if (v.length > 6) {
        input.value = `${v.slice(0, 6)}/${v.slice(6)}`;
    } else {
        input.value = v;
    }
}

function handleEnter(event) {
    if (event.key === 'Enter') {
        event.preventDefault();
        const form = document.getElementById('form-certificado');
        const inputs = Array.from(form.querySelectorAll('input'));
        const currentIndex = inputs.indexOf(document.activeElement);
        const nextIndex = (currentIndex + 1) % inputs.length;
        inputs[nextIndex].focus();
    }
}

function adicionarCertificado() {
    const form = document.getElementById('form-certificado');
    const dados = new FormData(form);
    const cert = Object.fromEntries(dados.entries());
    if (!cert.barcode || !cert.data || !cert.num_certificado || !cert.tipo_instrumento) {
        alert('Por favor, preencha os campos obrigatórios: Barcode, Data, Número do Certificado e Tipo de Instrumento.');
        return;
    }
    certificados.push(cert);
    atualizarListaVisual();
    ['data', 'num_certificado', 'tag', 'sala', 'modelo', 'fabricante'].forEach(id => {
        document.getElementById(id).value = '';
    });
    document.getElementById('data').focus();
}

function atualizarListaVisual() {
    const lista = document.getElementById('lista-certificados');
    lista.innerHTML = '';
    certificados.forEach((cert, index) => {
        const item = document.createElement('li');
        item.textContent = `Cert: ${cert.num_certificado} - TAG: ${cert.tag || 'N/A'}`;
        lista.appendChild(item);
    });
    document.getElementById('batch_data').value = JSON.stringify(certificados);
}

document.getElementById('form-lote').addEventListener('submit', async function(event) {
    event.preventDefault();

    const form = event.target;
    const submitButton = form.querySelector('button[type="submit"]');
    const originalButtonText = submitButton.innerHTML;

    submitButton.disabled = true;
    submitButton.innerHTML = 'Gerando...';

    try {
        const formData = new FormData(form);
        const response = await fetch(form.action, {
            method: form.method,
            body: formData
        });

        if (!response.ok) {
            const errorMessage = await response.text();
            alert(`Erro do servidor: ${errorMessage}`);
        } else {
            // Lidar com o download do arquivo
            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            const contentDisposition = response.headers.get('content-disposition');
            let fileName = 'lote.xlsx';
            if (contentDisposition) {
                const fileNameMatch = contentDisposition.match(/filename="(.+)"/);
                if (fileNameMatch && fileNameMatch.length > 1) {
                    fileName = fileNameMatch[1];
                }
            }
            a.href = downloadUrl;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(downloadUrl);
        }
    } catch (error) {
        console.error('Erro ao enviar o formulário:', error);
        alert('Ocorreu um erro de rede. Verifique sua conexão e tente novamente.');
    } finally {
        submitButton.disabled = false;
        submitButton.innerHTML = originalButtonText;
    }
});
