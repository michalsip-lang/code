async function loadAnalytics() {
    const brand = {
        blue: '#292982',
        grey: '#808184',
        red: '#e01b37',
        blueSoft: 'rgba(41, 41, 130, 0.16)',
        greySoft: 'rgba(128, 129, 132, 0.20)',
        redSoft: 'rgba(224, 27, 55, 0.18)'
    };

    const response = await fetch('/api/analytics/series');
    if (!response.ok) {
        return;
    }

    const data = await response.json();

    const dailyLabels = Object.keys(data.daily);
    const dailyValues = Object.values(data.daily);

    const weeklyLabels = Object.keys(data.weekly);
    const weeklyValues = Object.values(data.weekly);

    const monthlyLabels = Object.keys(data.monthly);
    const monthlyValues = Object.values(data.monthly);

    const stateLabels = data.states.map(s => s.status);
    const stateValues = data.states.map(s => s.count);

    new Chart(document.getElementById('dailyChart'), {
        type: 'line',
        data: {
            labels: dailyLabels,
            datasets: [{
                label: 'Dokončené úkoly / den',
                data: dailyValues,
                borderColor: brand.blue,
                backgroundColor: brand.blueSoft,
                tension: 0.3,
                fill: true
            }]
        }
    });

    new Chart(document.getElementById('weeklyChart'), {
        type: 'bar',
        data: {
            labels: weeklyLabels,
            datasets: [{
                label: 'Dokončené úkoly / týden',
                data: weeklyValues,
                backgroundColor: brand.red
            }]
        }
    });

    new Chart(document.getElementById('monthlyChart'), {
        type: 'line',
        data: {
            labels: monthlyLabels,
            datasets: [{
                label: 'Dokončené úkoly / měsíc',
                data: monthlyValues,
                borderColor: brand.grey,
                backgroundColor: brand.greySoft,
                tension: 0.3
            }]
        }
    });

    new Chart(document.getElementById('statusChart'), {
        type: 'doughnut',
        data: {
            labels: stateLabels,
            datasets: [{
                data: stateValues,
                backgroundColor: [brand.blue, brand.grey, brand.red, '#b8bbc4']
            }]
        }
    });

    renderHeatmap(data.heatmap);
}

function renderHeatmap(heatmap) {
    const host = document.getElementById('heatmap');
    if (!host) {
        return;
    }

    const labels = Object.keys(heatmap).sort();
    labels.forEach(label => {
        const val = heatmap[label];
        const div = document.createElement('div');
        const level = Math.min(4, val);
        div.className = `heat-cell heat-${level}`;
        div.title = `${label}: ${val}`;
        host.appendChild(div);
    });
}

loadAnalytics();
