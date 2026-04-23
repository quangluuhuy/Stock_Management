let data = JSON.parse(localStorage.getItem("stockout")) || [];

function render() {
  const table = document.getElementById("table");
  table.innerHTML = "";

  data.forEach(item => {
    table.innerHTML += `
      <tr>
        <td>${item.part}</td>
        <td>${item.serial}</td>
        <td>${item.desc}</td>
        <td>${item.status}</td>
        <td>${item.owner}</td>
        <td>${item.contract}</td>
        <td>${item.borrow}</td>
        <td>${item.date}</td>
      </tr>
    `;
  });
}

function addItem() {
  const part = document.getElementById("part").value;
  const serial = document.getElementById("serial").value;
  const desc = document.getElementById("desc").value;
  const status = document.getElementById("status").value;
  const owner = document.getElementById("owner").value;
  const contract = document.getElementById("contract").value;
  const borrow = document.getElementById("borrow").value;
  const date = document.getElementById("date").value;

  data.push({ part, serial, desc, status, owner, contract, borrow, date });
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
