document.addEventListener('DOMContentLoaded', () => {
    let certificados = [];
    const form = document.getElementById('form-certificado');
    const addButton = document.getElementById('add-btn');
    const editIndexField = document.getElementById('edit-index');
    const listaUI = document.getElementById('lista-certificados');
    const batchDataInput = document.getElementById('batch_data');

    document.getElementById('data').addEventListener('input', (e) => formatarData(e.target));
    document.getElementById('num_certificado').addEventListener('input', (e) => formatarCertificado(e.target));
    addButton.addEventListener('click', adicionarOuAtualizarCertificado);

    window.editarCertificado = editarCertificado;
    window.excluirCertificado = excluirCertificado;
    window.handleEnter = handleEnter;

    function handleEnter(event) {
        if (event.key === 'Enter') {
            event.preventDefault();
            const inputs = Array.from(form.querySelectorAll('input:not([type="hidden"])'));
            const currentIndex = inputs.indexOf(document.activeElement);
            if (currentIndex > -1 && currentIndex < inputs.length - 1) {
                inputs[currentIndex + 1].focus();
            } else if (currentIndex === inputs.length - 1) {
                addButton.click();
            }
        }
    }

    function formatarData(input) { let v = input.value.replace(/\D/g, '').slice(0, 8); if (v.length >= 5) { input.value = `${v.slice(0, 2)}/${v.slice(2, 4)}/${v.slice(4)}`; } else if (v.length >= 3) { input.value = `${v.slice(0, 2)}/${v.slice(2)}`; } else { input.value = v; } }
    function formatarCertificado(input) { let v = input.value.replace(/\D/g, '').slice(0, 8); if (v.length > 6) { input.value = `${v.slice(0, 6)}/${v.slice(6)}`; } else { input.value = v; } }
    
    // --- INÍCIO DA NOVA LÓGICA DE VALIDAÇÃO ---
    function validarData(dataStr) {
        // Expressão regular para verificar o formato dd/mm/aaaa
        const regex = /^\d{2}\/\d{2}\/\d{4}$/;
        if (!regex.test(dataStr)) {
            return false;
        }
        
        // Verifica se a data é válida (ex: não aceita 31/02/2024)
        const [dia, mes, ano] = dataStr.split('/').map(Number);
        const data = new Date(ano, mes - 1, dia); // Mês é 0-indexado
        return data.getFullYear() === ano && data.getMonth() === mes - 1 && data.getDate() === dia;
    }
    // --- FIM DA NOVA LÓGICA DE VALIDAÇÃO ---

    function adicionarOuAtualizarCertificado() {
        const dados = new FormData(form);
        const cert = Object.fromEntries(dados.entries());

        // --- VALIDAÇÃO NO FRONTEND ---
        if (!cert.barcode || !cert.data || !cert.num_certificado || !cert.tipo_instrumento) {
            alert('Por favor, preencha os campos obrigatórios.');
            return;
        }

        if (!validarData(cert.data)) {
            alert("Formato de data inválido. Por favor, use dd/mm/aaaa.");
            document.getElementById('data').focus(); // Foca no campo de data para correção
            return; // Interrompe a função aqui
        }
        // --- FIM DA VALIDAÇÃO ---

        const editIndex = parseInt(editIndexField.value, 10);
        if (editIndex > -1) {
            certificados[editIndex] = cert;
        } else {
            certificados.push(cert);
        }
        
        resetarFormulario();
        atualizarListaVisual();
    }

    function editarCertificado(index) {
        const cert = certificados[index];
        for (const key in cert) {
            if (form.elements[key]) {
                form.elements[key].value = cert[key];
            }
        }
        editIndexField.value = index;
        addButton.textContent = '💾 Atualizar Item';
        addButton.style.backgroundColor = '#ffc107';
        document.getElementById('data').focus();
    }

    function excluirCertificado(index) {
        if (confirm('Tem certeza que deseja excluir este item?')) {
            certificados.splice(index, 1);
            atualizarListaVisual();
            if (parseInt(editIndexField.value, 10) === index) {
                resetarFormulario();
            }
        }
    }

    function atualizarListaVisual() {
        listaUI.innerHTML = '';
        certificados.forEach((cert, index) => {
            const item = document.createElement('li');
            item.innerHTML = `
                <span>Cert: ${cert.num_certificado} - TAG: ${cert.tag || 'N/A'}</span>
                <div class="list-actions">
                    <span onclick="editarCertificado(${index})">✏️</span>
                    <span onclick="excluirCertificado(${index})">🗑️</span>
                </div>
            `;
            listaUI.appendChild(item);
        });
        batchDataInput.value = JSON.stringify(certificados);
    }

    function resetarFormulario() {
        const barcodeValue = form.elements['barcode'].value;
        form.reset();
        form.elements['barcode'].value = barcodeValue;
        editIndexField.value = -1;
        addButton.textContent = '+ Adicionar à Lista';
        addButton.style.backgroundColor = 'var(--btn-add)';
        document.getElementById('data').focus();
    }
});