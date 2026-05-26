---
layout: post
title: "Gender analysis of Waste Value Chain in Chattogram, Cox’s Bazar & Tangail"
date: 2025-10-20
description: Interactive xls form and Kobo questionnaire form
tags: kobo odk consultancy
categories: kobo
thumbnail: assets/img/9.jpg
---

You can read the abstract from here:  [👀📖✨ Read](/projects/12_project)



## Xls form

Use the tabs below to navigate between sheets of the dataset.

<div id="excel-container" style="width:100%; margin-bottom: 2rem;">
  <div id="sheet-tabs" style="display:flex; flex-wrap:wrap; gap:6px; margin-bottom:10px;"></div>
  <div style="overflow-x:auto;">
    <table id="excel-table" style="border-collapse:collapse; width:100%; font-size:0.85rem;"></table>
  </div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
(function() {
  const FILE_URL = '/assets/files/survey_house_undp.xlsx';

  let workbook;

  fetch(FILE_URL)
    .then(res => res.arrayBuffer())
    .then(buf => {
      workbook = XLSX.read(buf, { type: 'array' });
      const tabContainer = document.getElementById('sheet-tabs');

      workbook.SheetNames.forEach((name, index) => {
        const btn = document.createElement('button');
        btn.textContent = name;
        btn.style.cssText = 'padding:6px 14px; border:1px solid #ccc; border-radius:4px; cursor:pointer; background:' + (index === 0 ? 'var(--global-theme-color)' : '#f5f5f5') + '; color:' + (index === 0 ? '#fff' : '#333') + '; font-size:0.85rem;';
        btn.onclick = function() {
          document.querySelectorAll('#sheet-tabs button').forEach(b => {
            b.style.background = '#f5f5f5';
            b.style.color = '#333';
          });
          btn.style.background = 'var(--global-theme-color)';
          btn.style.color = '#fff';
          renderSheet(name);
        };
        tabContainer.appendChild(btn);
      });

      renderSheet(workbook.SheetNames[0]);
    })
    .catch(() => {
      document.getElementById('excel-container').innerHTML = '<p style="color:red;">Could not load the file. Please check the file path.</p>';
    });

  function renderSheet(sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const html = XLSX.utils.sheet_to_html(sheet, { editable: false });
    const table = document.getElementById('excel-table');
    const wrapper = document.createElement('div');
    wrapper.innerHTML = html;
    const newTable = wrapper.querySelector('table');
    if (newTable) {
      newTable.style.cssText = 'border-collapse:collapse; width:100%; font-size:0.85rem;';
      newTable.querySelectorAll('td, th').forEach(cell => {
        cell.style.cssText = 'border:1px solid #ddd; padding:6px 10px; text-align:left; white-space:nowrap;';
      });
      newTable.querySelectorAll('tr:first-child td, tr:first-child th').forEach(cell => {
        cell.style.background = '#f0f0f0';
        cell.style.fontWeight = '600';
      });
      table.parentNode.replaceChild(newTable, table);
      newTable.id = 'excel-table';
    }
  }
})();
</script>



## Survey Questionnaire

You can preview the survey form below:

<div style="margin-top: 1rem;">
  <iframe 
    src="https://ee.kobotoolbox.org/x/vkZiFUQt"
    width="100%" 
    height="800px"
    style="border:1px solid #ddd; border-radius:6px;"
    allowfullscreen>
  </iframe>
</div>

<p style="font-size:0.85rem; color:#888; margin-top:0.5rem;">
  If the form doesn't load above, 
  <a href="https://ee.kobotoolbox.org/x/vkZiFUQt" target="_blank">click here to open it directly</a>.
</p>
