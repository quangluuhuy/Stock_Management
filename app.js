let data = JSON.parse(localStorage.getItem("stockout")) || [];

function render() {
  const table = document.getElementById("table");
  table.innerHTML = "";

  data.forEach(item => {
    table.innerHTML += `
      <tr>
        <td>${item.part}</td>
        <td>${item.qty}</td>
        <td>${item.date}</td>
      </tr>
    `;
  });
}

function addItem() {
  const part = document.getElementById("part").value;
  const qty = document.getElementById("qty").value;
  const date = document.getElementById("date").value;

  data.push({ part, qty, date });
  localStorage.setItem("stockout", JSON.stringify(data));

  render();
}

function exportExcel() {
  const month = document.getElementById("month").value;

  const filtered = data.filter(item => item.date.startsWith(month));

  const ws = XLSX.utils.json_to_sheet(filtered);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Stockout");

  XLSX.writeFile(wb, `stockout_${month}.xlsx`);
}

render();
