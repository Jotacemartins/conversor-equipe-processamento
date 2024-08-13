document.getElementById('fileInput').addEventListener('change', () => {
    document.getElementById('downloadLink').style.display = 'none';
    document.getElementById('loadingMessage').style.display = 'none';
});

document.getElementById('convertButton').addEventListener('click', () => {
    const fileInput = document.getElementById('fileInput').files[0];
    const formatType = document.getElementById('formatSelect').value;

    if (!fileInput) {
        alert('Por favor, selecione um arquivo!');
        return;
    }

    document.getElementById('loadingMessage').style.display = 'block';

    const reader = new FileReader();

    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: '', raw: true});

        const updatedRows = rows.map((row, index) => {
            if (index === 0) {
                return formatType === 'cpf' ? ['CPF'] : ['Cartão'];
            }

            let cardNumber = String(row[0]);

            if (cardNumber.includes('E')) {
                cardNumber = Number(cardNumber).toFixed(0);
            }

            cardNumber = cardNumber.replace(/[^0-9]/g, '');

            if (formatType === "cartao1") {
                cardNumber = cardNumber.padStart(13, '0');
                return [formatCardNumber1(cardNumber)];
            } else if (formatType === "cpf") {
                cardNumber = cardNumber.padStart(11, '0');
                return [formatCPF(cardNumber)];
            } else if (formatType === "cartao2") {
                cardNumber = cardNumber.padStart(10, '0');
                return [formatCardNumber2(cardNumber)];
            }

            return row;
        });

        const newSheet = XLSX.utils.aoa_to_sheet(updatedRows);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);

        const wbout = XLSX.write(newWorkbook, {bookType: 'xlsx', type: 'array'});
        const blob = new Blob([wbout], {type: "application/octet-stream"});
        const url = URL.createObjectURL(blob);

        const downloadLink = document.getElementById('downloadLink');
        downloadLink.href = url;
        downloadLink.download = 'numeros_convertidos.xlsx';
        downloadLink.textContent = 'Baixar Números Convertidos';
        downloadLink.style.display = 'block';

        document.getElementById('loadingMessage').style.display = 'none';
    };

    reader.readAsArrayBuffer(fileInput);
});

function formatCardNumber1(number) {
    return `${number.substring(0, 2)}.${number.substring(2, 4)}.${number.substring(4, 12)}-${number.substring(12)}`;
}

function formatCPF(number) {
    return `${number.substring(0, 3)}.${number.substring(3, 6)}.${number.substring(6, 9)}-${number.substring(9)}`;
}

function formatCardNumber2(number) {
    return `${number.substring(0, 1)}.${number.substring(1, 4)}.${number.substring(4, 7)}.${number.substring(7, 10)}`;
}
