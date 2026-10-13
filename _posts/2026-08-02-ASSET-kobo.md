---
layout: post
title: "Independent Verification of Enterprise Based Training (EBT)"
date: 2026-08-02
description: "Interactive Excel forms and KoboToolbox questionnaires for trainees, institutions, and employers."
tags: kobo odk consultancy
categories: kobo
thumbnail: assets/img/9.jpg
---

You can read the abstract from here: [👀📖✨ Read](/projects/13_project)

This page provides separate Excel previews and KoboToolbox questionnaires for three respondent groups: trainees and graduates, institutions, and employers.

<!-- ================================================== -->
<!-- SHARED STYLES: EXCEL PREVIEWS                      -->
<!-- ================================================== -->

<style>
.excel-container {
  width: 100%;
  margin: 1rem 0 2rem;
}

.sheet-tabs {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-bottom: 10px;
}

.sheet-tabs button {
  padding: 6px 14px;
  border: 1px solid #ccc;
  border-radius: 4px;
  cursor: pointer;
  background: #f5f5f5;
  color: #333;
  font-size: 0.85rem;
  transition: background 0.2s ease;
}

.sheet-tabs button:focus-visible {
  outline: 2px solid var(--global-theme-color);
  outline-offset: 2px;
}

.sheet-tabs button.active {
  background: var(--global-theme-color);
  color: #fff;
}

.excel-table-wrapper {
  width: 100%;
  overflow-x: auto;
  -webkit-overflow-scrolling: touch;
}

.excel-table {
  border-collapse: collapse;
  width: 100%;
  font-size: 0.85rem;
}

.excel-table td,
.excel-table th {
  border: 1px solid #ddd;
  padding: 6px 10px;
  text-align: left;
  white-space: nowrap;
}

.excel-table tr:first-child td,
.excel-table tr:first-child th {
  background: #f0f0f0;
  font-weight: 600;
}

.excel-loading {
  padding: 12px;
  color: #666;
}

.excel-error {
  padding: 12px;
  color: #b42318;
  border: 1px solid #f1b5b0;
  border-radius: 4px;
}
</style>

<!-- ================================================== -->
<!-- 1. TRAINEES AND GRADUATES                           -->
<!-- ================================================== -->

## Excel Form: For Trainees and Graduates

The first three tabs (survey, choices, and settings) belong to the trainees and graduates questionnaire.

Use the tabs below to navigate between sheets.

<div id="excel-trainees" class="excel-container">
  <div class="sheet-tabs" aria-label="Trainees and graduates Excel sheets"></div>
  <div class="excel-table-wrapper" tabindex="0" aria-label="Trainees and graduates spreadsheet preview">
    <p class="excel-loading">Loading Excel file...</p>
  </div>
</div>

## Survey Questionnaire: For Trainees and Graduates

You can preview the survey form below.

<div style="margin-top: 1rem;">
  <iframe
    src="https://ee.kobotoolbox.org/x/lVMQ8q9f"
    title="KoboToolbox questionnaire for trainees and graduates"
    width="100%"
    height="800"
    style="border:1px solid #ddd; border-radius:6px;"
    loading="lazy"
    allowfullscreen>
  </iframe>
</div>

<p style="font-size:0.85rem; color:#888; margin-top:0.5rem;">
  If the form does not load above,
  <a href="https://ee.kobotoolbox.org/x/lVMQ8q9f" target="_blank" rel="noopener noreferrer">click here to open it directly</a>.
</p>

<!-- ================================================== -->
<!-- 2. INSTITUTIONS                                     -->
<!-- ================================================== -->

## Excel Form: For Institutions

Use the tabs below to navigate between sheets.

<div id="excel-institutions" class="excel-container">
  <div class="sheet-tabs" aria-label="Institutions Excel sheets"></div>
  <div class="excel-table-wrapper" tabindex="0" aria-label="Institutions spreadsheet preview">
    <p class="excel-loading">Loading Excel file...</p>
  </div>
</div>

## Survey Questionnaire: For Institutions

You can preview the survey form below.

<div style="margin-top: 1rem;">
  <iframe
    src="https://ee.kobotoolbox.org/x/KTKUrYnb"
    title="KoboToolbox questionnaire for institutions"
    width="100%"
    height="800"
    style="border:1px solid #ddd; border-radius:6px;"
    loading="lazy"
    allowfullscreen>
  </iframe>
</div>

<p style="font-size:0.85rem; color:#888; margin-top:0.5rem;">
  If the form does not load above,
  <a href="https://ee.kobotoolbox.org/x/KTKUrYnb" target="_blank" rel="noopener noreferrer">click here to open it directly</a>.
</p>

<!-- ================================================== -->
<!-- 3. EMPLOYERS                                        -->
<!-- ================================================== -->

## Excel Form: For Employers

Use the tabs below to navigate between sheets.

<div id="excel-employers" class="excel-container">
  <div class="sheet-tabs" aria-label="Employers Excel sheets"></div>
  <div class="excel-table-wrapper" tabindex="0" aria-label="Employers spreadsheet preview">
    <p class="excel-loading">Loading Excel file...</p>
  </div>
</div>

## Survey Questionnaire: For Employers

You can preview the survey form below.

<div style="margin-top: 1rem;">
  <iframe
    src="https://ee.kobotoolbox.org/x/037f19Y7"
    title="KoboToolbox questionnaire for employers"
    width="100%"
    height="800"
    style="border:1px solid #ddd; border-radius:6px;"
    loading="lazy"
    allowfullscreen>
  </iframe>
</div>

<p style="font-size:0.85rem; color:#888; margin-top:0.5rem;">
  If the form does not load above,
  <a href="https://ee.kobotoolbox.org/x/037f19Y7" target="_blank" rel="noopener noreferrer">click here to open it directly</a>.
</p>

<!-- ================================================== -->
<!-- SHARED JAVASCRIPT: LOAD THREE SEPARATE WORKBOOKS    -->
<!-- ================================================== -->

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<script>
(function () {
  'use strict';

  // Each workbook has its own container and file path.
  const excelFiles = [
    {
      containerId: 'excel-trainees',
      fileUrl: '/assets/files/EBT_Trainees%20and%20Graduates.xlsx',
      label: 'Trainees and Graduates'
    },
    {
      containerId: 'excel-institutions',
      fileUrl: '/assets/files/EBT_Institution%20Checklist.xlsx',
      label: 'Institutions'
    },
    {
      containerId: 'excel-employers',
      fileUrl: '/assets/files/EBT_Employer.xlsx',
      label: 'Employers'
    }
  ];

  function showError(container, message) {
    container.innerHTML = '';

    const error = document.createElement('p');
    error.className = 'excel-error';
    error.textContent = message;

    container.appendChild(error);
  }

  function loadExcel(config) {
    const container = document.getElementById(config.containerId);

    if (!container) {
      console.error(
        'Excel container not found:',
        config.containerId
      );
      return;
    }

    const tabs = container.querySelector('.sheet-tabs');
    const tableWrapper = container.querySelector(
      '.excel-table-wrapper'
    );

    if (!tabs || !tableWrapper) {
      showError(container, 'The Excel preview container is incomplete.');
      return;
    }

    if (typeof XLSX === 'undefined') {
      showError(
        container,
        'The Excel library could not be loaded. Please refresh the page and try again.'
      );
      return;
    }

    fetch(config.fileUrl)
      .then(function (response) {
        if (!response.ok) {
          throw new Error(
            'File request failed with HTTP status ' + response.status
          );
        }

        return response.arrayBuffer();
      })
      .then(function (buffer) {
        const workbook = XLSX.read(buffer, {
          type: 'array'
        });

        if (!workbook.SheetNames.length) {
          throw new Error('The workbook contains no worksheets.');
        }

        // Clear only this workbook's sheet tabs.
        tabs.replaceChildren();

        function renderSheet(sheetName, activeButton) {
          const sheet = workbook.Sheets[sheetName];

          if (!sheet) {
            tableWrapper.textContent =
              'The selected worksheet could not be found.';
            return;
          }

          // Render the selected sheet only inside this workbook.
          const html = XLSX.utils.sheet_to_html(sheet, {
            editable: false
          });

          const temporaryContainer = document.createElement('div');
          temporaryContainer.innerHTML = html;

          const renderedTable =
            temporaryContainer.querySelector('table');

          if (!renderedTable) {
            tableWrapper.textContent =
              'This worksheet could not be displayed.';
            return;
          }

          renderedTable.className = 'excel-table';
          renderedTable.setAttribute(
            'aria-label',
            config.label + ': ' + sheetName
          );

          // Replace the current table inside this preview only.
          tableWrapper.replaceChildren(renderedTable);

          // Update active tab styling inside this preview only.
          tabs.querySelectorAll('button').forEach(function (button) {
            button.classList.remove('active');
            button.setAttribute('aria-pressed', 'false');
          });

          activeButton.classList.add('active');
          activeButton.setAttribute('aria-pressed', 'true');
        }

        // Create an independent button for every worksheet.
        workbook.SheetNames.forEach(function (sheetName, index) {
          const button = document.createElement('button');

          button.type = 'button';
          button.textContent = sheetName;
          button.setAttribute('aria-pressed', 'false');

          button.addEventListener('click', function () {
            renderSheet(sheetName, button);
          });

          tabs.appendChild(button);

          // Display the first worksheet by default.
          if (index === 0) {
            renderSheet(sheetName, button);
          }
        });
      })
      .catch(function (error) {
        console.error(
          'Failed to load ' + config.label + ' Excel file:',
          error
        );

        showError(
          container,
          'Could not load the ' + config.label +
          ' Excel file. Please check the file path and filename.'
        );
      });
  }

  // Load all three workbooks independently.
  excelFiles.forEach(function (config) {
    loadExcel(config);
  });
})();
</script>
