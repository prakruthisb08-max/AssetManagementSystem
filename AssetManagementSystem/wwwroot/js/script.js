function normalizeMacField() {
    const input = document.querySelector('input[name="mac_address"]');
    if (!input || input.readOnly) return;
    input.value = input.value.replace(/-/g, ':').replace(/\s/g, '').toLowerCase();
}

function validateAssetForm() {
    normalizeMacField();

    const mac = document.querySelector('input[name="mac_address"]');
    if (mac && mac.value.trim() !== '') {
        const pattern = /^([0-9a-f]{2}:){5}[0-9a-f]{2}$/i;
        if (!pattern.test(mac.value.trim())) {
            alert('Enter a valid MAC Address, for example 00:11:22:33:44:55.');
            mac.focus();
            return false;
        }
    }

    const hd = document.querySelector('input[name="hard_disk"]');
    const used = document.querySelector('input[name="hard_disk_used"]');

    if (hd && used && hd.value !== '' && used.value !== '') {
        const total = parseFloat(hd.value);
        const usedValue = parseFloat(used.value);

        if (!isNaN(total) && !isNaN(usedValue) && usedValue > total) {
            alert('Hard Disk Used cannot be greater than Hard Disk capacity.');
            used.focus();
            return false;
        }
    }

    if (!validateNewAssetFields()) return false;

    return true;
}

function validateNewAssetFields() {
    const requiredSelects = ['zone','location','hostname','domainname','operating_system',
        'year_of_manufacture','network_type','warranty','staff_id','cost_center'];
    for (const name of requiredSelects) {
        const el = document.querySelector('[name="' + name + '"]');
        if (el && el.value.trim() === '') {
            alert('Please enter/select ' + name.replace(/_/g, ' ') + '.');
            el.focus();
            return false;
        }
    }
    return true;
}
