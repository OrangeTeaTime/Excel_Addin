// Assign event handlers and other initialization logic.
const headerList = document.getElementById('headers');
const tableList = document.getElementById('tables');

document.getElementById('get-tables').onclick = () => tryCatch(getTables);
document.getElementById('get-headers').onclick = () => tryCatch(getHeaders);
document.getElementById("apply-filters").onclick = () => tryCatch(customFilter);

async function getTables() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const tables = sheet.tables.load('name');
    await context.sync();

    tableList.innerHTML = '';

    optionElements = tables.items.map(table => {
      const option = document.createElement('option');
      option.value = table.name;
      option.textContent = table.name;
      return option;
    })
    
    tableList.append(...optionElements);
    await context.sync();
  });
}

async function getHeaders() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selected = document.getElementById('tables').value;
    const table = sheet.tables.getItem(selected);

    const headerRow = table.getHeaderRowRange().load('values');
    await context.sync();
    const headers = headerRow.values[0];

    const headerElements = headers.map(header => {
      const li = document.createElement('li');
      const label = document.createElement('label');
      const input = document.createElement('input');
      input.type = 'text';
      label.for = header;
      label.textContent = header;
      input.id = header;
      input.classList.add('filter');
      li.append(label, input);
      return li;
    });

    headerList.innerHTML = '';
    headerList.append(...headerElements);

    await context.sync();
  });
}

async function customFilter() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selected = document.getElementById('tables').value;
    const table = sheet.tables.getItem(selected);
    const filters = document.querySelectorAll('.filter');

    table.clearFilters();

    filters.forEach((column) => {
      if (column.value) {
        table.columns.getItem(column.id).filter.apply({
          criterion1: column.value,
          filterOn: Excel.FilterOn.custom,
        })
      };
    })

    await context.sync();
  })
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

