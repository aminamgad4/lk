// Enhanced Content script for ETA Invoice Exporter - Fixed Pagination Calculation
class ETAContentScript {
  constructor() {
    this.invoiceData = [];
    this.allPagesData = [];
    this.totalCount = 0;
    this.currentPage = 1;
    this.totalPages = 1;
    this.resultsPerPage = 50;
    this.isProcessingAllPages = false;
    this.progressCallback = null;
    this.domObserver = null;
    this.pageLoadTimeout = 20000; // 20 seconds timeout

    this.init();
  }
  
  init() {
    console.log('ETA Exporter: Content script initialized');
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => this.scanForInvoices());
    } else {
      setTimeout(() => this.scanForInvoices(), 1000);
    }
    
    this.setupMutationObserver();
  }
  
  setupMutationObserver() {
    this.observer = new MutationObserver((mutations) => {
      let shouldRescan = false;
      
      mutations.forEach((mutation) => {
        if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              if (node.classList?.contains('ms-DetailsRow') || 
                  node.querySelector?.('.ms-DetailsRow') ||
                  node.classList?.contains('ms-List-cell')) {
                shouldRescan = true;
              }
            }
          });
        }
      });
      
      if (shouldRescan && !this.isProcessingAllPages) {
        clearTimeout(this.rescanTimeout);
        this.rescanTimeout = setTimeout(() => this.scanForInvoices(), 800);
      }
    });
    
    this.observer.observe(document.body, {
      childList: true,
      subtree: true
    });
  }
  
  scanForInvoices() {
    try {
      console.log('ETA Exporter: Starting invoice scan...');
      this.invoiceData = [];
      
      // Extract pagination info first with improved logic
      this.extractPaginationInfoImproved();
      
      // Find invoice rows
      const rows = this.getVisibleInvoiceRows();
      console.log(`ETA Exporter: Found ${rows.length} visible invoice rows on page ${this.currentPage}`);
      
      if (rows.length === 0) {
        console.warn('ETA Exporter: No invoice rows found. Trying alternative selectors...');
        const alternativeRows = this.getAlternativeInvoiceRows();
        console.log(`ETA Exporter: Found ${alternativeRows.length} rows with alternative selectors`);
        alternativeRows.forEach((row, index) => {
          const invoiceData = this.extractDataFromRow(row, index + 1);
          if (this.isValidInvoiceData(invoiceData)) {
            this.invoiceData.push(invoiceData);
          }
        });
      } else {
        rows.forEach((row, index) => {
          const invoiceData = this.extractDataFromRow(row, index + 1);
          if (this.isValidInvoiceData(invoiceData)) {
            this.invoiceData.push(invoiceData);
          }
        });
      }
      
      // Update pagination info after scanning rows
      this.updatePaginationWithActualData();
      
      console.log(`ETA Exporter: Successfully extracted ${this.invoiceData.length} valid invoices from page ${this.currentPage}`);
      
    } catch (error) {
      console.error('ETA Exporter: Error scanning for invoices:', error);
    }
  }
  
  extractPaginationInfoImproved() {
    try {
      // Extract total count first
      this.totalCount = this.extractTotalCountImproved();
      
      // Extract current page
      this.currentPage = this.extractCurrentPageImproved();
      
      // Try to detect actual results per page from the UI
      this.resultsPerPage = this.detectResultsPerPageImproved();
      
      // Calculate total pages with better logic
      this.totalPages = this.calculateTotalPagesImproved();
      
      console.log(`ETA Exporter: Page ${this.currentPage} of ${this.totalPages}, Total: ${this.totalCount} invoices (${this.resultsPerPage} per page)`);
      
    } catch (error) {
      console.warn('ETA Exporter: Error extracting pagination info:', error);
      // Set safe defaults
      this.currentPage = Math.max(this.currentPage, 1);
      this.totalPages = Math.max(this.totalPages, this.currentPage);
      this.totalCount = Math.max(this.totalCount, 0);
    }
  }
  
  extractTotalCountImproved() {
    // Method 1: Look for "Results: XXX" text at bottom of page
    const resultElements = document.querySelectorAll('*');
    
    for (const element of resultElements) {
      const text = element.textContent || '';
      
      // Look for exact "Results: XXX" pattern
      const resultsMatch = text.match(/Results:\s*(\d+)/i);
      if (resultsMatch) {
        const count = parseInt(resultsMatch[1]);
        if (count > 0) {
          console.log(`ETA Exporter: Found total count ${count} from "Results:" text`);
          return count;
        }
      }
    }
    
    // Method 2: Look for pagination text like "1-50 of 304" or "Showing 1 to 50 of 304"
    const paginationPatterns = [
      /(\d+)\s*-\s*(\d+)\s*of\s*(\d+)/i,
      /showing\s*(\d+)\s*to\s*(\d+)\s*of\s*(\d+)/i,
      /(\d+)\s*-\s*(\d+)\s*من\s*(\d+)/i,
      /عرض\s*(\d+)\s*إلى\s*(\d+)\s*من\s*(\d+)/i,
      /صفحة\s*\d+\s*من\s*\d+\s*\(\s*(\d+)\s*عنصر\)/i
    ];
    
    for (const element of resultElements) {
      const text = element.textContent || '';
      
      for (const pattern of paginationPatterns) {
        const match = text.match(pattern);
        if (match) {
          const total = parseInt(match[3] || match[1]);
          if (total > 0) {
            console.log(`ETA Exporter: Found total count ${total} from pagination text`);
            return total;
          }
        }
      }
    }
    
    // Method 3: Look in specific pagination areas
    const paginationAreas = document.querySelectorAll(
      '.ms-CommandBar, [class*="pagination"], [class*="pager"], .ms-List-surface, [role="navigation"]'
    );
    
    for (const area of paginationAreas) {
      const text = area.textContent || '';
      
      // Look for total count indicators
      const totalMatch = text.match(/total:\s*(\d+)|النتائج:\s*(\d+)|(\d+)\s*results/i);
      if (totalMatch) {
        const total = parseInt(totalMatch[1] || totalMatch[2] || totalMatch[3]);
        if (total > 0) {
          console.log(`ETA Exporter: Found total count ${total} from pagination area`);
          return total;
        }
      }
    }
    
    // Method 4: Calculate from visible page numbers
    const maxVisiblePage = this.findMaxVisiblePageNumber();
    const estimatedPerPage = this.detectResultsPerPageImproved();
    
    if (maxVisiblePage > 0 && estimatedPerPage > 0) {
      const estimatedTotal = maxVisiblePage * estimatedPerPage;
      console.log(`ETA Exporter: Estimated total count ${estimatedTotal} from max page ${maxVisiblePage} * ${estimatedPerPage} per page`);
      return estimatedTotal;
    }
    
    return this.totalCount || 0;
  }
  
  extractCurrentPageImproved() {
    // Method 1: Look for active/selected page button with multiple indicators
    const activeSelectors = [
      'button[aria-pressed="true"]',
      'button[aria-current="page"]',
      '.ms-Button--primary[aria-pressed="true"]',
      'button.active',
      'button.selected',
      'button.current',
      'a[aria-current="page"]',
      'a.active',
      '.pagination .active',
      '.pager .active'
    ];
    
    for (const selector of activeSelectors) {
      const activeElement = document.querySelector(selector);
      if (activeElement) {
        const pageText = activeElement.textContent?.trim();
        const pageNum = parseInt(pageText);
        if (!isNaN(pageNum) && pageNum > 0) {
          console.log(`ETA Exporter: Found current page ${pageNum} from active element`);
          return pageNum;
        }
      }
    }
    
    // Method 2: Look for highlighted/styled page buttons
    const pageButtons = document.querySelectorAll('button, a');
    for (const button of pageButtons) {
      const text = button.textContent?.trim();
      const pageNum = parseInt(text);
      
      if (!isNaN(pageNum) && pageNum > 0) {
        // Check various indicators that this is the current page
        const computedStyle = window.getComputedStyle(button);
        const classes = button.className || '';
        const ariaPressed = button.getAttribute('aria-pressed');
        const ariaCurrent = button.getAttribute('aria-current');
        
        const isActive = (
          ariaPressed === 'true' ||
          ariaCurrent === 'page' ||
          classes.includes('active') ||
          classes.includes('selected') ||
          classes.includes('current') ||
          classes.includes('primary') ||
          computedStyle.backgroundColor !== 'rgba(0, 0, 0, 0)' ||
          computedStyle.fontWeight === 'bold' ||
          computedStyle.fontWeight === '700'
        );
        
        if (isActive) {
          console.log(`ETA Exporter: Found current page ${pageNum} from styled button`);
          return pageNum;
        }
      }
    }
    
    // Method 3: Check URL for page parameter
    const urlParams = new URLSearchParams(window.location.search);
    const urlPage = urlParams.get('page') || urlParams.get('p');
    if (urlPage) {
      const pageNum = parseInt(urlPage);
      if (!isNaN(pageNum) && pageNum > 0) {
        console.log(`ETA Exporter: Found current page ${pageNum} from URL`);
        return pageNum;
      }
    }
    
    return this.currentPage || 1;
  }
  
  detectResultsPerPageImproved() {
    // Method 1: Count actual visible rows (most reliable)
    const currentRowCount = this.invoiceData.length;
    console.log(`ETA Exporter: Current page has ${currentRowCount} rows`);
    
    // Method 2: Look for page size selector or indicator
    const pageSizeSelectors = [
      'select[aria-label*="page size"]',
      'select[aria-label*="items per page"]',
      '.ms-Dropdown-title',
      '[data-automation-id="pageSize"]'
    ];
    
    for (const selector of pageSizeSelectors) {
      const element = document.querySelector(selector);
      if (element) {
        const text = element.textContent || element.value || '';
        const sizeMatch = text.match(/(\d+)/);
        if (sizeMatch) {
          const pageSize = parseInt(sizeMatch[1]);
          if (pageSize > 0 && pageSize <= 200) {
            console.log(`ETA Exporter: Found page size ${pageSize} from selector`);
            return pageSize;
          }
        }
      }
    }
    
    // Method 3: Analyze pagination text to determine page size
    const paginationText = this.getPaginationText();
    if (paginationText) {
      // Look for patterns like "1-50 of 304" or "Showing 1 to 50 of 304"
      const rangeMatch = paginationText.match(/(\d+)\s*-\s*(\d+)\s*(?:of|من)\s*(\d+)/i);
      if (rangeMatch) {
        const start = parseInt(rangeMatch[1]);
        const end = parseInt(rangeMatch[2]);
        const pageSize = end - start + 1;
        if (pageSize > 0 && pageSize <= 200) {
          console.log(`ETA Exporter: Found page size ${pageSize} from pagination range`);
          return pageSize;
        }
      }
    }
    
    // Method 4: Use current row count if it makes sense
    if (currentRowCount > 0) {
      // Check if current row count matches common page sizes
      const commonSizes = [10, 20, 25, 30, 50, 100];
      
      // If we're not on the last page, use the current row count
      if (this.currentPage < this.totalPages) {
        // Check if it's close to a common page size
        for (const size of commonSizes) {
          if (Math.abs(currentRowCount - size) <= 2) {
            console.log(`ETA Exporter: Using page size ${size} (close to current ${currentRowCount})`);
            return size;
          }
        }
        
        // If not close to common sizes, use actual count
        if (currentRowCount <= 200) {
          console.log(`ETA Exporter: Using actual row count ${currentRowCount} as page size`);
          return currentRowCount;
        }
      }
      
      // If we're on the last page, estimate based on total count and page number
      if (this.totalCount > 0 && this.currentPage > 1) {
        const estimatedPageSize = Math.ceil(this.totalCount / this.currentPage);
        for (const size of commonSizes) {
          if (Math.abs(estimatedPageSize - size) <= 5) {
            console.log(`ETA Exporter: Estimated page size ${size} from total/pages calculation`);
            return size;
          }
        }
      }
    }
    
    // Method 5: Default based on ETA portal typical behavior
    console.log('ETA Exporter: Using default page size 50');
    return 50;
  }
  
  calculateTotalPagesImproved() {
    // Method 1: Find the highest visible page number
    const maxVisiblePage = this.findMaxVisiblePageNumber();
    
    // Method 2: Calculate from total count and results per page
    let calculatedPages = 1;
    if (this.totalCount > 0 && this.resultsPerPage > 0) {
      calculatedPages = Math.ceil(this.totalCount / this.resultsPerPage);
    }
    
    // Method 3: Look for "Page X of Y" text
    const pageOfMatch = this.getPaginationText().match(/page\s*(\d+)\s*of\s*(\d+)|صفحة\s*(\d+)\s*من\s*(\d+)/i);
    if (pageOfMatch) {
      const totalFromText = parseInt(pageOfMatch[2] || pageOfMatch[4]);
      if (totalFromText > 0) {
        console.log(`ETA Exporter: Found total pages ${totalFromText} from "page X of Y" text`);
        return totalFromText;
      }
    }
    
    // Choose the most reliable method
    const candidates = [maxVisiblePage, calculatedPages].filter(p => p > 0);
    
    if (candidates.length === 0) {
      return Math.max(this.currentPage, 1);
    }
    
    // Use the maximum of visible evidence
    const result = Math.max(...candidates, this.currentPage);
    
    console.log(`ETA Exporter: Calculated total pages: ${result} (max visible: ${maxVisiblePage}, calculated: ${calculatedPages})`);
    return result;
  }
  
  findMaxVisiblePageNumber() {
    const pageButtons = document.querySelectorAll('button, a');
    let maxPage = 0;
    
    pageButtons.forEach(btn => {
      const buttonText = btn.textContent?.trim();
      const pageNum = parseInt(buttonText);
      
      if (!isNaN(pageNum) && pageNum > maxPage && pageNum < 10000) { // Sanity check
        // Make sure this looks like a page button (not just any number)
        const parent = btn.parentElement;
        const classes = btn.className + ' ' + (parent?.className || '');
        
        if (classes.includes('page') || classes.includes('pagination') || 
            classes.includes('pager') || classes.includes('Button') ||
            btn.getAttribute('aria-label')?.includes('page')) {
          maxPage = pageNum;
        }
      }
    });
    
    return maxPage;
  }
  
  getPaginationText() {
    const paginationAreas = document.querySelectorAll(
      '.ms-CommandBar, [class*="pagination"], [class*="pager"], .ms-List-surface, [role="navigation"], footer'
    );
    
    let allText = '';
    paginationAreas.forEach(area => {
      allText += ' ' + (area.textContent || '');
    });
    
    return allText;
  }
  
  updatePaginationWithActualData() {
    // Update results per page based on actual data if needed
    const actualRowCount = this.invoiceData.length;
    
    if (actualRowCount > 0) {
      // If we're not on the last page and have a different count than expected, update
      if (this.currentPage < this.totalPages && actualRowCount !== this.resultsPerPage) {
        console.log(`ETA Exporter: Updating results per page from ${this.resultsPerPage} to ${actualRowCount}`);
        this.resultsPerPage = actualRowCount;
        
        // Recalculate total pages
        if (this.totalCount > 0) {
          this.totalPages = Math.ceil(this.totalCount / this.resultsPerPage);
          console.log(`ETA Exporter: Recalculated total pages: ${this.totalPages}`);
        }
      }
    }
  }
  
  getVisibleInvoiceRows() {
    // Primary selectors for invoice rows
    const selectors = [
      '.ms-DetailsRow[role="row"]',
      '.ms-List-cell[role="gridcell"]',
      '[data-list-index]',
      '.ms-DetailsRow',
      '[role="row"]'
    ];
    
    for (const selector of selectors) {
      const rows = document.querySelectorAll(selector);
      const visibleRows = Array.from(rows).filter(row => 
        this.isRowVisible(row) && this.hasInvoiceData(row)
      );
      
      if (visibleRows.length > 0) {
        console.log(`ETA Exporter: Found ${visibleRows.length} rows using selector: ${selector}`);
        return visibleRows;
      }
    }
    
    return [];
  }
  
  getAlternativeInvoiceRows() {
    // Alternative selectors when primary ones fail
    const alternativeSelectors = [
      'tr[role="row"]',
      '.ms-List-cell',
      '[data-automation-key]',
      '.ms-DetailsRow-cell',
      'div[role="gridcell"]'
    ];
    
    const allRows = [];
    
    for (const selector of alternativeSelectors) {
      const elements = document.querySelectorAll(selector);
      Array.from(elements).forEach(element => {
        const row = element.closest('[role="row"]') || element.parentElement;
        if (row && this.hasInvoiceData(row) && !allRows.includes(row)) {
          allRows.push(row);
        }
      });
    }
    
    return allRows.filter(row => this.isRowVisible(row));
  }
  
  isRowVisible(row) {
    if (!row) return false;
    
    const rect = row.getBoundingClientRect();
    const style = window.getComputedStyle(row);
    
    return (
      rect.width > 0 && 
      rect.height > 0 &&
      style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      style.opacity !== '0'
    );
  }
  
  hasInvoiceData(row) {
    if (!row) return false;
    
    // Check for electronic number or internal number
    const electronicNumber = row.querySelector('.internalId-link a, [data-automation-key="uuid"] a, .griCellTitle');
    const internalNumber = row.querySelector('.griCellSubTitle, [data-automation-key="uuid"] .griCellSubTitle');
    const totalAmount = row.querySelector('[data-automation-key="total"], .griCellTitleGray');
    
    return !!(electronicNumber?.textContent?.trim() || 
              internalNumber?.textContent?.trim() || 
              totalAmount?.textContent?.trim());
  }
  
  extractDataFromRow(row, index) {
    const invoice = {
      index: index,
      pageNumber: this.currentPage,
      
      // Main invoice data matching Excel format
      serialNumber: index,
      viewButton: 'عرض',
      documentType: 'فاتورة',
      documentVersion: '1.0',
      status: '',
      issueDate: '',
      submissionDate: '',
      invoiceCurrency: 'EGP',
      invoiceValue: '',
      vatAmount: '',
      taxDiscount: '0',
      totalInvoice: '',
      internalNumber: '',
      electronicNumber: '',
      sellerTaxNumber: '',
      sellerName: '',
      sellerAddress: '',
      buyerTaxNumber: '',
      buyerName: '',
      buyerAddress: '',
      purchaseOrderRef: '',
      purchaseOrderDesc: '',
      salesOrderRef: '',
      electronicSignature: 'موقع إلكترونياً',
      foodDrugGuide: '',
      externalLink: '',
      
      // Additional fields
      issueTime: '',
      totalAmount: '',
      currency: 'EGP',
      submissionId: '',
      details: []
    };
    
    try {
      // Try to extract data using different methods
      this.extractUsingDataAttributes(row, invoice);
      this.extractUsingCellPositions(row, invoice);
      this.extractUsingTextContent(row, invoice);
      
      // Generate external link if we have electronic number
      if (invoice.electronicNumber) {
        invoice.externalLink = this.generateExternalLink(invoice);
      }
      
      // Note: Addresses will be extracted later when navigating to detail pages
      // This is handled in the getAllPagesData method when downloadDetails is enabled
      
    } catch (error) {
      console.warn(`ETA Exporter: Error extracting data from row ${index}:`, error);
    }
    
    return invoice;
  }
  
  extractUsingDataAttributes(row, invoice) {
    // Method 1: Using data-automation-key attributes
    const cells = row.querySelectorAll('.ms-DetailsRow-cell, [data-automation-key]');
    
    cells.forEach(cell => {
      const key = cell.getAttribute('data-automation-key');
      
      switch (key) {
        case 'uuid':
          const electronicLink = cell.querySelector('.internalId-link a.griCellTitle, a');
          if (electronicLink) {
            invoice.electronicNumber = electronicLink.textContent?.trim() || '';
          }
          
          const internalNumberElement = cell.querySelector('.griCellSubTitle');
          if (internalNumberElement) {
            invoice.internalNumber = internalNumberElement.textContent?.trim() || '';
          }
          break;
          
        case 'dateTimeReceived':
          const dateElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const timeElement = cell.querySelector('.griCellSubTitle');
          
          if (dateElement) {
            invoice.issueDate = dateElement.textContent?.trim() || '';
            invoice.submissionDate = invoice.issueDate;
          }
          if (timeElement) {
            invoice.issueTime = timeElement.textContent?.trim() || '';
          }
          break;
          
        case 'typeName':
          const typeElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const versionElement = cell.querySelector('.griCellSubTitle');
          
          if (typeElement) {
            invoice.documentType = typeElement.textContent?.trim() || 'فاتورة';
          }
          if (versionElement) {
            invoice.documentVersion = versionElement.textContent?.trim() || '1.0';
          }
          break;
          
        case 'total':
          const totalElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          if (totalElement) {
            const totalText = totalElement.textContent?.trim() || '';
            invoice.totalAmount = totalText;
            invoice.totalInvoice = totalText;
            
            // Calculate VAT and invoice value
            const totalValue = this.parseAmount(totalText);
            if (totalValue > 0) {
              const vatRate = 0.14;
              const vatAmount = (totalValue * vatRate) / (1 + vatRate);
              const invoiceValue = totalValue - vatAmount;
              
              invoice.vatAmount = this.formatAmount(vatAmount);
              invoice.invoiceValue = this.formatAmount(invoiceValue);
            }
          }
          break;
          
        case 'issuerName':
          const sellerNameElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const sellerTaxElement = cell.querySelector('.griCellSubTitle');
          
          if (sellerNameElement) {
            invoice.sellerName = sellerNameElement.textContent?.trim() || '';
          }
          if (sellerTaxElement) {
            invoice.sellerTaxNumber = sellerTaxElement.textContent?.trim() || '';
          }
          
          if (invoice.sellerName && !invoice.sellerAddress) {
            invoice.sellerAddress = 'غير محدد';
          }
          break;
          
        case 'receiverName':
          const buyerNameElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const buyerTaxElement = cell.querySelector('.griCellSubTitle');
          
          if (buyerNameElement) {
            invoice.buyerName = buyerNameElement.textContent?.trim() || '';
          }
          if (buyerTaxElement) {
            invoice.buyerTaxNumber = buyerTaxElement.textContent?.trim() || '';
          }
          
          if (invoice.buyerName && !invoice.buyerAddress) {
            invoice.buyerAddress = 'غير محدد';
          }
          break;
          
        case 'submission':
          const submissionLink = cell.querySelector('a.submissionId-link, a');
          if (submissionLink) {
            invoice.submissionId = submissionLink.textContent?.trim() || '';
            invoice.purchaseOrderRef = invoice.submissionId;
          }
          break;
          
        case 'status':
          const validRejectedDiv = cell.querySelector('.horizontal.valid-rejected');
          if (validRejectedDiv) {
            const validStatus = validRejectedDiv.querySelector('.status-Valid');
            const rejectedStatus = validRejectedDiv.querySelector('.status-Rejected');
            if (validStatus && rejectedStatus) {
              invoice.status = `${validStatus.textContent?.trim()} → ${rejectedStatus.textContent?.trim()}`;
            }
          } else {
            const textStatus = cell.querySelector('.textStatus, .griCellTitle, .griCellTitleGray');
            if (textStatus) {
              invoice.status = textStatus.textContent?.trim() || '';
            }
          }
          break;
      }
    });
  }
  
  extractUsingCellPositions(row, invoice) {
    // Method 2: Using cell positions (fallback)
    const cells = row.querySelectorAll('.ms-DetailsRow-cell, td, [role="gridcell"]');
    
    if (cells.length >= 8) {
      // Try to extract based on typical column positions
      if (!invoice.electronicNumber) {
        const firstCell = cells[0];
        const link = firstCell.querySelector('a');
        if (link) {
          invoice.electronicNumber = link.textContent?.trim() || '';
        }
      }
      
      if (!invoice.totalAmount) {
        // Total amount is usually in one of the middle columns
        for (let i = 2; i < Math.min(6, cells.length); i++) {
          const cellText = cells[i].textContent?.trim() || '';
          if (cellText.includes('EGP') || /^\d+[\d,]*\.?\d*$/.test(cellText.replace(/[,٬]/g, ''))) {
            invoice.totalAmount = cellText;
            invoice.totalInvoice = cellText;
            break;
          }
        }
      }
      
      if (!invoice.issueDate) {
        // Date is usually in one of the early columns
        for (let i = 1; i < Math.min(4, cells.length); i++) {
          const cellText = cells[i].textContent?.trim() || '';
          if (cellText.includes('/') && cellText.length >= 8) {
            invoice.issueDate = cellText;
            invoice.submissionDate = cellText;
            break;
          }
        }
      }
    }
  }
  
  extractUsingTextContent(row, invoice) {
    // Method 3: Extract from all text content (last resort)
    const allText = row.textContent || '';
    
    // Extract electronic number pattern
    if (!invoice.electronicNumber) {
      const electronicMatch = allText.match(/[A-Z0-9]{20,30}/);
      if (electronicMatch) {
        invoice.electronicNumber = electronicMatch[0];
      }
    }
    
    // Extract date pattern
    if (!invoice.issueDate) {
      const dateMatch = allText.match(/\d{1,2}\/\d{1,2}\/\d{4}/);
      if (dateMatch) {
        invoice.issueDate = dateMatch[0];
        invoice.submissionDate = dateMatch[0];
      }
    }
    
    // Extract amount pattern
    if (!invoice.totalAmount) {
      const amountMatch = allText.match(/\d+[,٬]?\d*\.?\d*\s*EGP/);
      if (amountMatch) {
        invoice.totalAmount = amountMatch[0];
        invoice.totalInvoice = amountMatch[0];
      }
    }
  }
  
  parseAmount(amountText) {
    if (!amountText) return 0;
    const cleanText = amountText.replace(/[,٬\sEGP]/g, '').replace(/[^\d.]/g, '');
    return parseFloat(cleanText) || 0;
  }
  
  formatAmount(amount) {
    if (!amount || amount === 0) return '0';
    return amount.toLocaleString('en-US', { 
      minimumFractionDigits: 2, 
      maximumFractionDigits: 2 
    });
  }
  
  generateExternalLink(invoice) {
    if (!invoice.electronicNumber) return '';
    
    let shareId = '';
    if (invoice.submissionId && invoice.submissionId.length > 10) {
      shareId = invoice.submissionId;
    } else {
      shareId = invoice.electronicNumber.replace(/[^A-Z0-9]/g, '').substring(0, 26);
    }
    
    return `https://invoicing.eta.gov.eg/documents/${invoice.electronicNumber}/share/${shareId}`;
  }
  
  isValidInvoiceData(invoice) {
    return !!(invoice.electronicNumber || invoice.internalNumber || invoice.totalAmount);
  }
  
  async getAllPagesData(options = {}) {
    const {
      downloadDetails = false,
      downloadAllPages = false,
      combineAll = false,
      progressCallback = null
    } = options;
    
    try {
      this.isProcessingAllPages = true;
      this.allPagesData = [];
      
      console.log(`ETA Exporter: Starting to load ALL pages. Total invoices to load: ${this.totalCount}, Total pages: ${this.totalPages}`);
      
      // Always start from page 1 to ensure we get everything
      await this.navigateToFirstPage();
      
      let processedInvoices = 0;
      let currentPageNum = 1;
      let maxAttempts = this.totalPages + 10; // Safety limit based on calculated pages
      let attempts = 0;
      
      // Continue until we've processed all invoices or reached limits
      while (processedInvoices < this.totalCount && attempts < maxAttempts && currentPageNum <= this.totalPages) {
        attempts++;
        
        try {
          console.log(`ETA Exporter: Processing page ${currentPageNum}, attempt ${attempts}...`);
          
          // Update progress
          if (this.progressCallback) {
            this.progressCallback({
              currentPage: currentPageNum,
              totalPages: this.totalPages,
              message: `جاري معالجة الصفحة ${currentPageNum}... (${processedInvoices}/${this.totalCount} فاتورة)`,
              percentage: this.totalCount > 0 ? (processedInvoices / this.totalCount) * 100 : 0
            });
          }
          
          // Wait for page to load completely
          await this.waitForPageLoadComplete();
          
          // Scan invoices on current page
          this.scanForInvoices();
          
          if (this.invoiceData.length > 0) {
            // Add page data to collection with correct serial numbers
            const pageData = this.invoiceData.map((invoice, index) => ({
              ...invoice,
              pageNumber: currentPageNum,
              serialNumber: processedInvoices + index + 1,
              globalIndex: processedInvoices + index + 1
            }));
            
            // Extract addresses if downloadDetails is enabled OR if address fields are selected
            const shouldExtractAddresses = downloadDetails || 
              (options.sellerAddress && options.sellerAddress !== undefined) || 
              (options.buyerAddress && options.buyerAddress !== undefined);
            
            if (shouldExtractAddresses) {
              console.log('ETA Exporter: Extracting addresses for invoices on this page...');
              
              for (let i = 0; i < pageData.length; i++) {
                const invoice = pageData[i];
                
                if (progressCallback) {
                  progressCallback({
                    type: 'detail',
                    message: `جاري استخراج عناوين الفاتورة ${i + 1} من ${pageData.length} في الصفحة ${currentPageNum}`,
                    current: i + 1,
                    total: pageData.length,
                    page: currentPageNum,
                    totalPages: this.totalPages
                  });
                }
                
                try {
                  // Extract addresses from invoice detail page
                  const addressData = await this.extractAddressesFromInvoiceDetail(invoice);
                  if (addressData) {
                    // Only update addresses if the corresponding option is selected
                    if (downloadDetails || (options.sellerAddress && options.sellerAddress !== undefined)) {
                      invoice.sellerAddress = addressData.sellerAddress || 'غير محدد';
                    }
                    if (downloadDetails || (options.buyerAddress && options.buyerAddress !== undefined)) {
                      invoice.buyerAddress = addressData.buyerAddress || 'غير محدد';
                    }
                  }
                } catch (error) {
                  console.warn(`ETA Exporter: Error extracting addresses for invoice ${invoice.electronicNumber}:`, error);
                }
                
                // Small delay between requests
                await this.delay(1000);
              }
            }
            
            this.allPagesData.push(...pageData);
            processedInvoices += this.invoiceData.length;
            
            console.log(`ETA Exporter: Page ${currentPageNum} processed, collected ${this.invoiceData.length} invoices. Total: ${processedInvoices}/${this.totalCount}`);
          } else {
            console.warn(`ETA Exporter: No invoices found on page ${currentPageNum}`);
          }
          
          // Check if we've got all invoices
          if (processedInvoices >= this.totalCount) {
            console.log(`ETA Exporter: Successfully loaded all ${processedInvoices} invoices!`);
            break;
          }
          
          // Check if we've reached the last page
          if (currentPageNum >= this.totalPages) {
            console.log(`ETA Exporter: Reached last page (${this.totalPages})`);
            break;
          }
          
          // Try to navigate to next page
          const navigatedToNext = await this.navigateToNextPageRobust();
          if (!navigatedToNext) {
            console.log('ETA Exporter: Cannot navigate to next page, stopping');
            break;
          }
          
          currentPageNum++;
          await this.delay(2000); // Wait between pages
          
        } catch (error) {
          console.error(`Error processing page ${currentPageNum}:`, error);
          
          // Try to continue to next page
          const navigatedToNext = await this.navigateToNextPageRobust();
          if (!navigatedToNext) {
            break;
          }
          currentPageNum++;
        }
      }
      
      console.log(`ETA Exporter: Completed! Loaded ${this.allPagesData.length} invoices out of ${this.totalCount} total.`);
      
      return {
        success: true,
        data: this.allPagesData,
        totalProcessed: this.allPagesData.length,
        expectedTotal: this.totalCount
      };
      
    } catch (error) {
      console.error('ETA Exporter: Error getting all pages data:', error);
      return { 
        success: false, 
        data: this.allPagesData,
        error: error.message,
        totalProcessed: this.allPagesData.length
      };
    } finally {
      this.isProcessingAllPages = false;
    }
  }
  
  async extractAddressesFromInvoiceDetail(invoice) {
    try {
      console.log(`ETA Exporter: Extracting addresses for invoice ${invoice.electronicNumber}...`);
      
      // Store current URL
      const currentUrl = window.location.href;
      
      // Navigate to invoice detail page
      const detailUrl = `/documents/${invoice.electronicNumber}`;
      window.location.href = detailUrl;
      
      // Wait for page to load
      await this.waitForPageLoadComplete();
      
      // Extract addresses from the detail page
      const addresses = this.extractAddressesFromCurrentPage();
      
      // Return to the list page
      await this.returnToMainList(currentUrl);
      
      console.log(`ETA Exporter: Successfully extracted addresses for invoice ${invoice.electronicNumber}:`, {
        sellerAddress: addresses.sellerAddress ? 'Found' : 'Not found',
        buyerAddress: addresses.buyerAddress ? 'Found' : 'Not found'
      });
      
      return addresses;
      
    } catch (error) {
      console.error(`ETA Exporter: Error extracting addresses for invoice ${invoice.electronicNumber}:`, error);
      
      // Try to return to the list page
      try {
        await this.returnToMainList(currentUrl);
      } catch (returnError) {
        console.error('ETA Exporter: Error returning to list page:', returnError);
      }
      
      return {
        sellerAddress: null,
        buyerAddress: null
      };
    }
  }
  
  async extractAddressesForSingleInvoice(invoiceId, options) {
    try {
      console.log(`ETA Exporter: Extracting addresses for single invoice ${invoiceId}...`);
      
      // Store current URL
      const currentUrl = window.location.href;
      
      // Navigate to invoice detail page
      const detailUrl = `/documents/${invoiceId}`;
      window.location.href = detailUrl;
      
      // Wait for page to load
      await this.waitForPageLoadComplete();
      
      // Extract addresses from the detail page
      const addresses = this.extractAddressesFromCurrentPage();
      
      // Return to the list page
      await this.returnToMainList(currentUrl);
      
      console.log(`ETA Exporter: Successfully extracted addresses for invoice ${invoiceId}:`, {
        sellerAddress: addresses.sellerAddress ? 'Found' : 'Not found',
        buyerAddress: addresses.buyerAddress ? 'Found' : 'Not found'
      });
      
      return {
        success: true,
        addresses: addresses
      };
      
    } catch (error) {
      console.error(`ETA Exporter: Error extracting addresses for invoice ${invoiceId}:`, error);
      
      // Try to return to the list page
      try {
        await this.returnToMainList(currentUrl);
      } catch (returnError) {
        console.error('ETA Exporter: Error returning to list page:', returnError);
      }
      
      return {
        success: false,
        addresses: {
          sellerAddress: null,
          buyerAddress: null
        },
        error: error.message
      };
    }
  }
  
  async navigateToFirstPage() {
    console.log('ETA Exporter: Navigating to first page...');
    
    // Look for page 1 button
    const pageButtons = document.querySelectorAll('button, a');
    for (const button of pageButtons) {
      const buttonText = button.textContent?.trim();
      if (buttonText === '1') {
        console.log('ETA Exporter: Clicking page 1 button');
        button.click();
        await this.delay(3000);
        await this.waitForPageLoadComplete();
        this.extractPaginationInfoImproved();
        return;
      }
    }
    
    // If no page 1 button, we might already be on page 1
    this.extractPaginationInfoImproved();
    if (this.currentPage === 1) {
      console.log('ETA Exporter: Already on page 1');
      return;
    }
    
    // Try to navigate using previous buttons until we reach page 1
    while (this.currentPage > 1) {
      const navigated = await this.navigateToPreviousPage();
      if (!navigated) break;
      await this.delay(2000);
      await this.waitForPageLoadComplete();
      this.extractPaginationInfoImproved();
    }
  }
  
  async navigateToNextPageRobust() {
    console.log('ETA Exporter: Attempting to navigate to next page...');
    
    // Method 1: Try to find and click next button
    const nextButton = this.findNextButton();
    if (nextButton && !nextButton.disabled) {
      console.log('ETA Exporter: Found next button, clicking...');
      nextButton.click();
      await this.delay(3000);
      return true;
    }
    
    // Method 2: Try to find the next page number button
    const currentPageNum = this.currentPage;
    const nextPageNum = currentPageNum + 1;
    
    const pageButtons = document.querySelectorAll('button, a');
    for (const button of pageButtons) {
      const buttonText = button.textContent?.trim();
      if (parseInt(buttonText) === nextPageNum) {
        console.log(`ETA Exporter: Found page ${nextPageNum} button, clicking...`);
        button.click();
        await this.delay(3000);
        return true;
      }
    }
    
    // Method 3: Try to find any button that might be "next"
    for (const button of pageButtons) {
      const buttonText = button.textContent?.toLowerCase().trim();
      const ariaLabel = button.getAttribute('aria-label')?.toLowerCase() || '';
      
      if (buttonText.includes('next') || buttonText.includes('التالي') || 
          buttonText === '>' || buttonText === '»' ||
          ariaLabel.includes('next') || ariaLabel.includes('التالي')) {
        
        if (!button.disabled && !button.getAttribute('disabled')) {
          console.log('ETA Exporter: Found potential next button, clicking...');
          button.click();
          await this.delay(3000);
          return true;
        }
      }
    }
    
    console.log('ETA Exporter: No next page navigation found');
    return false;
  }
  
  findNextButton() {
    const nextSelectors = [
      'button[aria-label*="Next"]',
      'button[aria-label*="next"]',
      'button[title*="Next"]',
      'button[title*="next"]',
      'button[aria-label*="التالي"]',
      'button:has([data-icon-name="ChevronRight"])',
      'button:has([class*="chevron-right"])',
      '.ms-Button:has([data-icon-name="ChevronRight"])',
      'button:has(i[class*="right"])',
      'button:has(span[class*="right"])'
    ];
    
    for (const selector of nextSelectors) {
      try {
        const button = document.querySelector(selector);
        if (button && !button.disabled && !button.getAttribute('disabled')) {
          const style = window.getComputedStyle(button);
          if (style.display !== 'none' && style.visibility !== 'hidden') {
            return button;
          }
        }
      } catch (e) {
        // Ignore selector errors
      }
    }
    
    // Fallback: look for buttons with right arrow icons or text
    const allButtons = document.querySelectorAll('button');
    for (const button of allButtons) {
      if (button.disabled || button.getAttribute('disabled')) continue;
      
      const hasRightArrow = button.querySelector('[data-icon-name="ChevronRight"], [class*="chevron-right"], [class*="arrow-right"]');
      if (hasRightArrow) {
        return button;
      }
      
      // Check button text for next indicators
      const text = button.textContent?.toLowerCase() || '';
      if (text.includes('next') || text.includes('التالي') || text === '>' || text === '»') {
        return button;
      }
    }
    
    return null;
  }
  
  async navigateToPreviousPage() {
    const prevSelectors = [
      'button[aria-label*="Previous"]',
      'button[aria-label*="previous"]',
      'button[title*="Previous"]',
      'button[title*="previous"]',
      'button[aria-label*="السابق"]',
      'button:has([data-icon-name="ChevronLeft"])',
      'button:has([class*="chevron-left"])',
      '.ms-Button:has([data-icon-name="ChevronLeft"])'
    ];
    
    for (const selector of prevSelectors) {
      try {
        const button = document.querySelector(selector);
        if (button && !button.disabled) {
          const style = window.getComputedStyle(button);
          if (style.display !== 'none' && style.visibility !== 'hidden') {
            console.log('ETA Exporter: Clicking previous button');
            button.click();
            await this.delay(2000);
            return true;
          }
        }
      } catch (e) {
        // Ignore selector errors
      }
    }
    
    console.warn('ETA Exporter: No previous button found');
    return false;
  }
  
  async waitForPageLoadComplete() {
    console.log('ETA Exporter: Waiting for page load to complete...');
    
    // Wait for loading indicators to disappear
    await this.waitForCondition(() => {
      const loadingIndicators = document.querySelectorAll(
        '.LoadingIndicator, .ms-Spinner, [class*="loading"], [class*="spinner"], .ms-Shimmer'
      );
      const isLoading = Array.from(loadingIndicators).some(el => 
        el.offsetParent !== null && 
        window.getComputedStyle(el).display !== 'none'
      );
      return !isLoading;
    }, 20000);
    
    // Wait for invoice rows to appear
    await this.waitForCondition(() => {
      const rows = this.getVisibleInvoiceRows();
      return rows.length > 0;
    }, 20000);
    
    // Wait for DOM stability
    await this.delay(3000);
    
    console.log('ETA Exporter: Page load completed');
  }
  
  async waitForCondition(condition, timeout = 15000) {
    const startTime = Date.now();
    
    while (Date.now() - startTime < timeout) {
      try {
        if (condition()) {
          return true;
        }
      } catch (error) {
        // Ignore errors in condition check
      }
      await this.delay(500);
    }
    
    console.warn(`ETA Exporter: Condition timeout after ${timeout}ms`);
    return false;
  }
  
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  
  setProgressCallback(callback) {
    this.progressCallback = callback;
  }
  
  async getInvoiceDetails(invoiceId) {
    try {
      const details = await this.extractInvoiceDetailsFromPage(invoiceId);
      
      // Extract addresses from the current page
      const addresses = this.extractAddressesFromCurrentPage();
      
      return {
        success: true,
        data: details,
        addresses: addresses
      };
    } catch (error) {
      console.error('Error getting invoice details:', error);
      return { 
        success: false, 
        data: [],
        addresses: {},
        error: error.message 
      };
    }
  }
  
  extractAddressesFromCurrentPage() {
    try {
      console.log('ETA Exporter: Extracting addresses from current page...');
      console.log('ETA Exporter: Current URL:', window.location.href);
      
      const addresses = {
        sellerAddress: null,
        buyerAddress: null
      };
      
      // Extract seller address
      addresses.sellerAddress = this.findAddressInPage('seller', 'البائع', 'seller', 'rtl');
      
      // Extract buyer address
      addresses.buyerAddress = this.findAddressInPage('buyer', 'المشتري', 'buyer', 'rtl');
      
      console.log('ETA Exporter: Address extraction results:', {
        sellerAddress: addresses.sellerAddress ? `Found: ${addresses.sellerAddress.substring(0, 50)}...` : 'Not found',
        buyerAddress: addresses.buyerAddress ? `Found: ${addresses.buyerAddress.substring(0, 50)}...` : 'Not found'
      });
      
      return addresses;
    } catch (error) {
      console.error('ETA Exporter: Error extracting addresses from current page:', error);
      return {
        sellerAddress: null,
        buyerAddress: null
      };
    }
  }
  
  async extractInvoiceDetailsFromPage(invoiceId) {
    const details = [];
    
    try {
      // Look for details table
      const detailsTable = document.querySelector('.ms-DetailsList, [data-automationid="DetailsList"], table');
      
      if (detailsTable) {
        const rows = detailsTable.querySelectorAll('.ms-DetailsRow[role="row"], tr');
        
        rows.forEach((row, index) => {
          const cells = row.querySelectorAll('.ms-DetailsRow-cell, td');
          
          if (cells.length >= 6) {
            const item = {
              itemCode: this.extractCellText(cells[0]) || `ITEM-${index + 1}`,
              description: this.extractCellText(cells[1]) || 'صنف',
              unitCode: this.extractCellText(cells[2]) || 'EA',
              unitName: this.extractCellText(cells[3]) || 'قطعة',
              quantity: this.extractCellText(cells[4]) || '1',
              unitPrice: this.extractCellText(cells[5]) || '0',
              totalValue: this.extractCellText(cells[6]) || '0',
              taxAmount: this.extractCellText(cells[7]) || '0',
              vatAmount: this.extractCellText(cells[8]) || '0'
            };
            
            // Skip header rows
            if (item.description && 
                item.description !== 'اسم الصنف' && 
                item.description !== 'Description' &&
                item.description.trim() !== '') {
              details.push(item);
            }
          }
        });
      }
      
      // If no details found, create a summary item
      if (details.length === 0) {
        const invoice = this.invoiceData.find(inv => inv.electronicNumber === invoiceId);
        if (invoice) {
          details.push({
            itemCode: invoice.electronicNumber || 'INVOICE',
            description: 'إجمالي الفاتورة',
            unitCode: 'EA',
            unitName: 'فاتورة',
            quantity: '1',
            unitPrice: invoice.totalAmount || '0',
            totalValue: invoice.invoiceValue || invoice.totalAmount || '0',
            taxAmount: '0',
            vatAmount: invoice.vatAmount || '0'
          });
        }
      }
      
    } catch (error) {
      console.error('Error extracting invoice details:', error);
    }
    
    return details;
  }
  
  extractCellText(cell) {
    if (!cell) return '';
    
    const textElement = cell.querySelector('.griCellTitle, .griCellTitleGray, .ms-DetailsRow-cellContent') || cell;
    return textElement.textContent?.trim() || '';
  }
  
  getInvoiceData() {
    return {
      invoices: this.invoiceData,
      totalCount: this.totalCount,
      currentPage: this.currentPage,
      totalPages: this.totalPages
    };
  }
  








  findAddressInPage(type, arabicType, englishType, direction) {
    try {
      console.log(`ETA Exporter: Looking for ${type} address...`);
      
      // Method 1: Look for specific textarea elements by ID (most reliable)
      if (type === 'seller') {
        const sellerAddressTextarea = document.querySelector('textarea[id="TextField378"]');
        if (sellerAddressTextarea) {
          const address = sellerAddressTextarea.value?.trim();
          if (address && address.length > 5) {
            console.log(`ETA Exporter: Found seller address from textarea: ${address.substring(0, 50)}...`);
            return address;
          }
        }
      } else if (type === 'buyer') {
        const buyerAddressTextarea = document.querySelector('textarea[id="TextField393"]');
        if (buyerAddressTextarea) {
          const address = buyerAddressTextarea.value?.trim();
          if (address && address.length > 5) {
            console.log(`ETA Exporter: Found buyer address from textarea: ${address.substring(0, 50)}...`);
            return address;
          }
        }
      }
      
      // Method 2: Look for address by label text and find the corresponding textarea
      const labels = document.querySelectorAll('label.ms-Label');
      for (const label of labels) {
        const text = label.textContent?.trim();
        if (text) {
          if ((type === 'seller' && (text.includes('Branch Address') || text.includes('عنوان الفرع'))) ||
              (type === 'buyer' && (text.includes('Address') && !text.includes('Branch')))) {
            
            // Find the parent container and look for textarea within it
            const container = label.closest('.vertialFlexDiv');
            if (container) {
              const textarea = container.querySelector('textarea');
              if (textarea) {
                const address = textarea.value?.trim();
                if (address && address.length > 5) {
                  console.log(`ETA Exporter: Found ${type} address from label "${text}": ${address.substring(0, 50)}...`);
                  return address;
                }
              }
            }
          }
        }
      }
      
      // Method 3: Look for address by class patterns
      if (type === 'seller') {
        const sellerAddressElements = document.querySelectorAll('.ms-TextField textarea');
        for (const element of sellerAddressElements) {
          const parent = element.closest('.vertialFlexDiv');
          if (parent && !parent.classList.contains('receiverAddress')) {
            const address = element.value?.trim();
            if (address && address.length > 5) {
              console.log(`ETA Exporter: Found seller address from class pattern: ${address.substring(0, 50)}...`);
              return address;
            }
          }
        }
      } else if (type === 'buyer') {
        const buyerAddressElements = document.querySelectorAll('.receiverAddress textarea');
        for (const element of buyerAddressElements) {
          const address = element.value?.trim();
          if (address && address.length > 5) {
            console.log(`ETA Exporter: Found buyer address from receiverAddress class: ${address.substring(0, 50)}...`);
            return address;
          }
        }
      }
      
      // Method 4: Fallback to original pattern matching
      let address = this.findAddressByPattern(type, arabicType, englishType);
      
      if (!address) {
        // Try to find by data attributes
        const elements = document.querySelectorAll('[data-address], [data-addr]');
        for (const element of elements) {
          const attr = element.getAttribute('data-address') || element.getAttribute('data-addr');
          if (attr && attr.toLowerCase().includes(type.toLowerCase())) {
            address = this.extractAddressFromElement(element);
            if (address) break;
          }
        }
      }
      
      if (!address) {
        console.log(`ETA Exporter: No ${type} address found on this page`);
      }
      
      return address;
    } catch (error) {
      console.error(`ETA Exporter: Error finding ${type} address:`, error);
      return null;
    }
  }

  findAddressByPattern(type, arabicType, englishType) {
    // Look for address patterns in the page text
    const pageText = document.body.textContent || '';
    
    // Common address patterns
    const patterns = [
      // Arabic patterns
      new RegExp(`${arabicType}[\\s\\n]*عنوان[\\s\\n]*:?[\\s\\n]*([^\\n\\r]+)`, 'i'),
      new RegExp(`عنوان[\\s\\n]*${arabicType}[\\s\\n]*:?[\\s\\n]*([^\\n\\r]+)`, 'i'),
      
      // English patterns
      new RegExp(`${englishType}[\\s\\n]*address[\\s\\n]*:?[\\s\\n]*([^\\n\\r]+)`, 'i'),
      new RegExp(`address[\\s\\n]*${englishType}[\\s\\n]*:?[\\s\\n]*([^\\n\\r]+)`, 'i'),
      
      // Generic patterns
      new RegExp(`${type}[\\s\\n]*address[\\s\\n]*:?[\\s\\n]*([^\\n\\r]+)`, 'i'),
      new RegExp(`address[\\s\\n]*${type}[\\s\\n]*:?[\\s\\n]*([^\\n\\r]+)`, 'i')
    ];

    for (const pattern of patterns) {
      const match = pageText.match(pattern);
      if (match && match[1]) {
        const address = match[1].trim();
        if (address && address.length > 5 && address !== 'غير محدد') {
          return address;
        }
      }
    }

    return null;
  }

  extractAddressFromElement(element) {
    if (!element) return null;

    // Try different methods to extract address
    const methods = [
      // Method 1: Direct text content
      () => element.textContent?.trim(),
      
      // Method 2: Value attribute
      () => element.value?.trim(),
      
      // Method 3: Placeholder attribute
      () => element.placeholder?.trim(),
      
      // Method 4: Title attribute
      () => element.title?.trim(),
      
      // Method 5: Data attributes
      () => element.getAttribute('data-value')?.trim(),
      
      // Method 6: Inner text of child elements
      () => {
        const textElements = element.querySelectorAll('span, div, p');
        for (const textEl of textElements) {
          const text = textEl.textContent?.trim();
          if (text && text.length > 5) {
            return text;
          }
        }
        return null;
      }
    ];

    for (const method of methods) {
      try {
        const address = method();
        if (address && address.length > 5 && address !== 'غير محدد') {
          return address;
        }
      } catch (error) {
        // Continue to next method
      }
    }

    return null;
  }

  async returnToMainList(originalUrl) {
    try {
      // Method 1: Try browser back button
      if (window.history.length > 1) {
        window.history.back();
        await this.delay(2000);
        
        // Check if we're back to the main list
        if (this.isMainInvoiceList()) {
          console.log('ETA Exporter: Successfully returned to main list via back button');
          return true;
        }
      }

      // Method 2: Navigate to original URL
      if (originalUrl && originalUrl !== window.location.href) {
        window.location.href = originalUrl;
        await this.delay(3000);
        
        if (this.isMainInvoiceList()) {
          console.log('ETA Exporter: Successfully returned to main list via URL navigation');
          return true;
        }
      }

      // Method 3: Try to find and click a "back to list" or "documents" link
      const backSelectors = [
        'a[href*="/documents"]',
        'a[href*="/invoices"]',
        'a[href*="/list"]',
        '.back-button',
        '.return-button',
        '[data-automationid*="back"]',
        '[data-automationid*="return"]'
      ];

      for (const selector of backSelectors) {
        const backLink = document.querySelector(selector);
        if (backLink) {
          backLink.click();
          await this.delay(2000);
          
          if (this.isMainInvoiceList()) {
            console.log('ETA Exporter: Successfully returned to main list via back link');
            return true;
          }
        }
      }

      console.warn('ETA Exporter: Could not return to main list, continuing anyway');
      return false;

    } catch (error) {
      console.error('ETA Exporter: Error returning to main list:', error);
      return false;
    }
  }

  isMainInvoiceList() {
    // Check for indicators that we're on the main invoice list page
    const indicators = [
      // URL patterns
      window.location.pathname.includes('/documents') && !window.location.pathname.includes('/documents/'),
      window.location.pathname === '/',
      
      // Page content indicators
      document.querySelector('.ms-DetailsList[data-automationid="DetailsList"]'),
      document.querySelector('.recentDocuments'),
      document.querySelector('[role="grid"]'),
      
      // Pagination indicators
      document.querySelector('.eta-pagination'),
      document.querySelector('[class*="pagination"]'),
      
      // Multiple invoice rows
      document.querySelectorAll('.ms-DetailsRow').length > 5
    ];
    
    return indicators.some(indicator => indicator);
  }
  
  cleanup() {
    if (this.observer) {
      this.observer.disconnect();
    }
    if (this.rescanTimeout) {
      clearTimeout(this.rescanTimeout);
    }
  }
}

// Initialize content script
const etaContentScript = new ETAContentScript();

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('ETA Exporter: Received message:', request.action);
  
  switch (request.action) {
    case 'ping':
      sendResponse({ success: true, message: 'Content script is ready' });
      break;
      
    case 'getInvoiceData':
      const data = etaContentScript.getInvoiceData();
      console.log('ETA Exporter: Returning invoice data:', data);
      sendResponse({
        success: true,
        data: data
      });
      break;
      
    case 'getInvoiceDetails':
      etaContentScript.getInvoiceDetails(request.invoiceId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;
      
    case 'getAllPagesData':
      if (request.options && request.options.progressCallback) {
        etaContentScript.setProgressCallback((progress) => {
          chrome.runtime.sendMessage({
            action: 'progressUpdate',
            progress: progress
          }).catch(() => {
            // Ignore errors if popup is closed
          });
        });
      }
      
      etaContentScript.getAllPagesData(request.options)
        .then(result => {
          console.log('ETA Exporter: All pages data result:', result);
          sendResponse(result);
        })
        .catch(error => {
          console.error('ETA Exporter: Error in getAllPagesData:', error);
          sendResponse({ success: false, error: error.message });
        });
      return true;
      
    case 'extractAddressesForInvoice':
      etaContentScript.extractAddressesForSingleInvoice(request.invoiceId, request.options)
        .then(result => {
          console.log('ETA Exporter: Address extraction result:', result);
          sendResponse(result);
        })
        .catch(error => {
          console.error('ETA Exporter: Error in extractAddressesForInvoice:', error);
          sendResponse({ success: false, error: error.message });
        });
      return true;
      
    case 'rescanPage':
      etaContentScript.scanForInvoices();
      sendResponse({
        success: true,
        data: etaContentScript.getInvoiceData()
      });
      break;
      
    default:
      sendResponse({ success: false, error: 'Unknown action' });
  }
  
  return true;
});

// Cleanup on page unload
window.addEventListener('beforeunload', () => {
  etaContentScript.cleanup();
});

console.log('ETA Exporter: Content script loaded successfully');