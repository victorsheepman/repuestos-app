const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyFZXk2dajXFg2FS5K1vFWltojWcigbTAO3-VSK7eCnWCoNszD5MACjpRUXR-zjdPF3/exec";

// Toast
const toastContainer = document.createElement('div');
toastContainer.className = 'fixed top-6 right-6 flex flex-col gap-2 z-50';
document.body.appendChild(toastContainer);

function showToast(message, type = 'success') {
    const colors = {
        success: 'bg-emerald-600',
        error:   'bg-red-600',
        info:    'bg-blue-600',
    };
    const icons = {
        success: 'fa-circle-check',
        error:   'fa-circle-xmark',
        info:    'fa-circle-info',
    };
    const toast = document.createElement('div');
    toast.className = `${colors[type]} text-white text-sm font-semibold px-4 py-3 rounded-xl shadow-lg flex items-center gap-3 animate-fade-in`;
    toast.innerHTML = `<i class="fa-solid ${icons[type]}"></i> <span>${message}</span>`;
    toastContainer.appendChild(toast);
    setTimeout(() => toast.remove(), 3500);
}

const SHEETS_CSV_URL = "https://docs.google.com/spreadsheets/d/1h8LGQ8ODEChZkWjo2w2VqFHtDtzmUUTqAyzXMs-z7YQ/export?format=csv&gid=0";

let currentData = [];

let selections = {
    tipo: null,
    zona: null,
    estacion: null,
    marcaModelo: null,
    repuestos: []
};

function getVal(row, keys) {
    for (let k of Object.keys(row)) {
        let cleanK = k.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
        for (let target of keys) {
            let cleanTarget = target.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
            if (cleanK === cleanTarget) return row[k];
        }
    }
    return '';
}

const resetBtn = document.getElementById('resetBtn');
const cartList = document.getElementById('cartList');
const selectedCount = document.getElementById('selectedCount');
const copyBtn = document.getElementById('copyBtn');
const whatsappBtn = document.getElementById('whatsappBtn');
const sheetsBtn = document.getElementById('sheetsBtn');
const pathSummary = document.getElementById('pathSummary');

// Carga de Excel dinámico
document.getElementById('excelUpload').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
        try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            if (jsonData.length > 0) {
                currentData = jsonData;
                document.getElementById('dataStatus').innerHTML = `
                    <i class="fa-solid fa-circle-check text-emerald-600 mr-2 text-base"></i>
                    <strong>¡Archivo cargado!</strong> Se leyeron <strong>${currentData.length} repuestos</strong> desde "${file.name}".
                `;
                resetBtn.click();
            } else {
                showToast('El archivo no contiene filas válidas.', 'error');
            }
        } catch (err) {
            showToast('Error al leer el archivo. Asegúrate de que sea un Excel o CSV válido.', 'error');
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
});

// Carga desde Google Sheets
async function loadData() {
    const dataStatus = document.getElementById('dataStatus');
    const overlay = document.getElementById('loadingOverlay');

    overlay.classList.remove('hidden');
    dataStatus.innerHTML = `<i class="fa-solid fa-spinner fa-spin text-blue-600 mr-2 text-base"></i> <strong>Cargando datos desde Google Sheets...</strong>`;
    document.getElementById('statusBanner').className = 'bg-blue-50 border border-blue-200 text-blue-800 p-4 rounded-2xl text-sm flex items-center justify-between shadow-sm';

    try {
        const res = await fetch(SHEETS_CSV_URL);
        if (!res.ok) throw new Error('No se pudo acceder al Sheet');
        const csv = await res.text();
        const workbook = XLSX.read(csv, { type: 'string' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        if (jsonData.length === 0) throw new Error('El Sheet está vacío');

        console.log('Datos cargados desde Google Sheets:', jsonData);
        currentData = jsonData;
        dataStatus.innerHTML = `<i class="fa-solid fa-circle-check text-emerald-600 mr-2 text-base"></i> <strong>¡Datos cargados!</strong> Se encontraron <strong>${currentData.length} repuestos</strong> desde Google Sheets.`;
        document.getElementById('statusBanner').className = 'bg-emerald-50 border border-emerald-200 text-emerald-800 p-4 rounded-2xl text-sm flex items-center justify-between shadow-sm';
    } catch (err) {
        console.error(err);
        dataStatus.innerHTML = `<i class="fa-solid fa-triangle-exclamation text-amber-600 mr-2 text-base"></i> <strong>No se pudo cargar el Sheet.</strong> Verifica tu conexión o que el Sheet sea público.`;
        document.getElementById('statusBanner').className = 'bg-amber-50 border border-amber-200 text-amber-800 p-4 rounded-2xl text-sm flex items-center justify-between shadow-sm';
    } finally {
        overlay.classList.add('hidden');
    }

    renderStep1();
}

// Inicialización
document.addEventListener('DOMContentLoaded', loadData);

resetBtn.addEventListener('click', () => {
    selections = { tipo: null, zona: null, estacion: null, marcaModelo: null, repuestos: [] };
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step3').classList.add('hidden');
    document.getElementById('step4').classList.add('hidden');
    document.getElementById('step5').classList.add('hidden');
    pathSummary.classList.add('hidden');
    renderStep1();
    updateSummary();
});

// Paso 1: Tipo
function renderStep1() {
    const container = document.getElementById('options1');
    container.innerHTML = '';
    const tipos = [...new Set(currentData.map(d => getVal(d, ['TIPO', 'TIPO_OPERACION', 'CATEGORIA'])))].filter(Boolean);

    tipos.forEach(tipo => {
        const btn = document.createElement('button');
        const isSelected = selections.tipo === tipo;
        const icon = tipo.toUpperCase().includes('ELECTR') ? 'fa-bolt' : 'fa-oil-can';

        btn.className = `p-4 rounded-xl border-2 text-left font-bold transition flex items-center gap-3 ${isSelected ? 'border-blue-600 bg-blue-50 text-blue-800' : 'border-slate-200 hover:border-blue-300 bg-white'}`;
        btn.innerHTML = `<i class="fa-solid ${icon} text-lg text-blue-600"></i> <span>${tipo}</span>`;
        btn.onclick = () => {
            selections.tipo = tipo;
            selections.zona = null; selections.estacion = null; selections.marcaModelo = null; selections.repuestos = [];
            renderStep1();
            renderStep2();
            updateSummary();
        };
        container.appendChild(btn);
    });
}

// Paso 2: Zona
function renderStep2() {
    const step2 = document.getElementById('step2');
    const container = document.getElementById('options2');
    step2.classList.remove('hidden');
    container.innerHTML = '';

    const filteredData = currentData.filter(d => getVal(d, ['TIPO']) === selections.tipo);
    const zonas = [...new Set(filteredData.map(d => getVal(d, ['ZONA', 'ESTADO', 'REGION'])))].filter(Boolean);

    zonas.forEach(zona => {
        const btn = document.createElement('button');
        const isSelected = selections.zona === zona;
        btn.className = `p-3 rounded-xl border-2 text-sm font-semibold transition ${isSelected ? 'border-blue-600 bg-blue-50 text-blue-800' : 'border-slate-200 hover:border-blue-300 bg-white'}`;
        btn.innerText = zona;
        btn.onclick = () => {
            selections.zona = zona;
            selections.estacion = null; selections.marcaModelo = null; selections.repuestos = [];
            renderStep2();
            renderStep3();
            updateSummary();
        };
        container.appendChild(btn);
    });
    document.getElementById('step3').classList.add('hidden');
    document.getElementById('step4').classList.add('hidden');
    document.getElementById('step5').classList.add('hidden');
}

// Paso 3: Estación
function renderStep3() {
    const step3 = document.getElementById('step3');
    const container = document.getElementById('options3');
    step3.classList.remove('hidden');
    container.innerHTML = '';

    const filteredData = currentData.filter(d => getVal(d, ['TIPO']) === selections.tipo && getVal(d, ['ZONA']) === selections.zona);
    const estaciones = [...new Set(filteredData.map(d => getVal(d, ['ESTACION', 'ESTACIÓN', 'NOMBRE_ESTACION'])))].filter(Boolean);

    estaciones.forEach(estacion => {
        const btn = document.createElement('button');
        const isSelected = selections.estacion === estacion;
        btn.className = `p-3 rounded-xl border-2 text-sm font-semibold text-left transition ${isSelected ? 'border-blue-600 bg-blue-50 text-blue-800' : 'border-slate-200 hover:border-blue-300 bg-white'}`;
        btn.innerText = estacion;
        btn.onclick = () => {
            selections.estacion = estacion;
            selections.repuestos = [];
            renderStep3();
            renderStep4();
            renderStep5();
            updateSummary();
        };
        container.appendChild(btn);
    });
    document.getElementById('step4').classList.add('hidden');
    document.getElementById('step5').classList.add('hidden');
}

// Paso 4: Marca / Modelo
function renderStep4() {
    const step4 = document.getElementById('step4');
    const container = document.getElementById('options4');
    step4.classList.remove('hidden');

    const item = currentData.find(d =>
        getVal(d, ['TIPO']) === selections.tipo &&
        getVal(d, ['ZONA']) === selections.zona &&
        getVal(d, ['ESTACION']) === selections.estacion
    );

    if (item) {
        const marca = getVal(item, ['MARCA']);
        const modelo = getVal(item, ['MODELO']);
        selections.marcaModelo = `${marca} ${modelo}`.trim();
        container.innerHTML = `
            <div class="flex items-center gap-3">
                <i class="fa-solid fa-dispenser text-blue-600 text-xl"></i>
                <div>
                    <p class="text-xs text-slate-500 uppercase font-bold">Marca & Modelo Detectado</p>
                    <p class="text-base font-extrabold text-slate-800">${selections.marcaModelo}</p>
                </div>
            </div>
        `;
    }
}

// Paso 5: Repuestos
function renderStep5() {
    const step5 = document.getElementById('step5');
    const container = document.getElementById('options5');
    step5.classList.remove('hidden');
    container.innerHTML = '';

    const repuestosDisponibles = currentData.filter(d =>
        getVal(d, ['TIPO']) === selections.tipo &&
        getVal(d, ['ZONA']) === selections.zona &&
        getVal(d, ['ESTACION']) === selections.estacion
    );

    repuestosDisponibles.forEach(item => {
        const nombreRepuesto = getVal(item, ['REPUESTO']);
        const codigo = getVal(item, ['CODIGO', 'NOMBRE']);
        const isChecked = selections.repuestos.some(r => r.nombre === codigo && r.repuesto === nombreRepuesto);

        const label = document.createElement('label');
        label.className = `flex items-center justify-between p-3.5 rounded-xl border-2 cursor-pointer transition ${isChecked ? 'border-blue-600 bg-blue-50' : 'border-slate-200 hover:border-slate-300 bg-white'}`;
        label.innerHTML = `
            <div class="flex items-center gap-3">
                <input type="checkbox" ${isChecked ? 'checked' : ''} class="w-4 h-4 text-blue-600 rounded focus:ring-blue-500">
                <div>
                    <p class="text-sm font-bold text-slate-800">${nombreRepuesto}</p>
                    <p class="text-xs text-slate-400">Código: ${codigo}</p>
                </div>
            </div>
            <span class="text-xs bg-slate-100 text-slate-600 px-2 py-1 rounded-md font-semibold">${getVal(item, ['MARCA'])}</span>
        `;

        label.querySelector('input').onchange = (e) => {
            if (e.target.checked) {
                selections.repuestos.push({ repuesto: nombreRepuesto, nombre: codigo });
            } else {
                selections.repuestos = selections.repuestos.filter(r => !(r.nombre === codigo && r.repuesto === nombreRepuesto));
            }
            renderStep5();
            updateSummary();
        };

        container.appendChild(label);
    });
}

// Actualizar Resumen Lateral
function updateSummary() {
    if (selections.tipo) {
        pathSummary.classList.remove('hidden');
        document.getElementById('sumTipo').innerText = selections.tipo || '-';
        document.getElementById('sumZona').innerText = selections.zona || '-';
        document.getElementById('sumEstacion').innerText = selections.estacion || '-';
        document.getElementById('sumEquipo').innerText = selections.marcaModelo || '-';
    }

    selectedCount.innerText = `${selections.repuestos.length} ítems`;

    if (selections.repuestos.length === 0) {
        cartList.innerHTML = `<p class="text-slate-400 text-center py-10 text-xs italic">Sigue los pasos de la izquierda para seleccionar los repuestos que necesitas.</p>`;
        copyBtn.disabled = true;
        whatsappBtn.disabled = true;
        sheetsBtn.disabled = true;
    } else {
        cartList.innerHTML = '';
        selections.repuestos.forEach((r, idx) => {
            const div = document.createElement('div');
            div.className = 'py-2.5 flex items-center justify-between gap-2';
            div.innerHTML = `
                <div>
                    <p class="font-bold text-slate-800 text-xs">${r.repuesto}</p>
                    <p class="text-[10px] text-slate-400">Código: ${r.nombre}</p>
                </div>
                <button onclick="removeItem(${idx})" class="text-slate-300 hover:text-red-500 transition text-xs"><i class="fa-solid fa-trash-can"></i></button>
            `;
            cartList.appendChild(div);
        });
        copyBtn.disabled = false;
        whatsappBtn.disabled = false;
        sheetsBtn.disabled = false;
    }
}

function removeItem(index) {
    selections.repuestos.splice(index, 1);
    if (document.getElementById('step5').classList.contains('hidden') === false) {
        renderStep5();
    }
    updateSummary();
}

// Guardar en Google Sheets vía API
sheetsBtn.addEventListener('click', async () => {
    if (selections.repuestos.length === 0) return;

    const originalBtnText = sheetsBtn.innerHTML;
    sheetsBtn.disabled = true;
    sheetsBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Guardando...`;

    const payload = {
        tipo: selections.tipo,
        zona: selections.zona,
        estacion: selections.estacion,
        marcaModelo: selections.marcaModelo,
        repuestos: selections.repuestos
    };

    try {
        await fetch(GOOGLE_SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        showToast('¡Requerimiento registrado en Google Sheets!', 'success');
    } catch (error) {
        showToast('Error al conectar con Google Sheets. Verifica la conexión.', 'error');
        console.error(error);
    } finally {
        sheetsBtn.disabled = false;
        sheetsBtn.innerHTML = originalBtnText;
    }
});

// Copiar Portapapeles
copyBtn.addEventListener('click', () => {
    let text = `REQUERIMIENTO DE REPUESTOS\n`;
    text += `Sistema: ${selections.tipo}\n`;
    text += `Ubicación: ${selections.zona} - ${selections.estacion}\n`;
    text += `Equipo: ${selections.marcaModelo}\n\n`;
    text += `REPUESTOS REQUERIDOS:\n`;
    selections.repuestos.forEach((r, i) => {
        text += `${i + 1}. ${r.repuesto} (${r.nombre})\n`;
    });

    navigator.clipboard.writeText(text);
    showToast('¡Lista copiada al portapapeles!', 'info');
});

// Enviar WhatsApp
whatsappBtn.addEventListener('click', () => {
    let text = `*REQUERIMIENTO DE REPUESTOS*\n`;
    text += `*Sistema:* ${selections.tipo}\n`;
    text += `*Ubicación:* ${selections.zona} - ${selections.estacion}\n`;
    text += `*Equipo:* ${selections.marcaModelo}\n\n`;
    text += `*REPUESTOS REQUERIDOS:*\n`;
    selections.repuestos.forEach((r) => {
        text += `• ${r.repuesto} _(${r.nombre})_\n`;
    });

    window.open(`https://api.whatsapp.com/send?text=${encodeURIComponent(text)}`, '_blank');
});
