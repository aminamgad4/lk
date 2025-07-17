// Content script for ETA Invoice Exporter
class ETAContentScript {
    constructor() {
        this.isInitialized = false;
        this.currentUrl = window.location.href;
        this.init();
    }

    init() {
        if (this.isInitialized) return;
        
        // Only initialize on ETA invoicing site
        if (!window.location.href.includes('invoicing.eta.gov.eg')) {
            return;
        }

        this.isInitialized = true;
        this.setupMessageListener();
        this.addIndicator();
        
        console.log('ETA Invoice Exporter: Content script initialized');
    }

    setupMessageListener() {
        chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
            this.handleMessage(message, sender, sendResponse);
            return true; // Keep message channel open for async responses
        });
    }

    async handleMessage(message, sender, sendResponse) {
        try {
            switch (message.action) {
                case 'ping':
                    sendResponse({ success: true, message: 'Content script ready' });
                    break;

                case 'getInvoiceData':
                    const invoiceData = await this.extractInvoiceData();
                    sendResponse(invoiceData);
                    break;

                case 'getAllPagesData':
                    const allData = await this.extractAllPagesData(message.options);
                    sendResponse(allData);
                    break;

                case 'getInvoiceDetails':
                    const details = await this.extractInvoiceDetails(message.invoiceId);
                    sendResponse(details);
                    break;

                case 'extractAddressesForInvoice':
                    const addresses = await this.extractAddressesForInvoice(message.invoiceId, message.options);
                    sendResponse(addresses);
                    break;

                default:
                    sendResponse({ success: false, error: 'Unknown action' });
            }
        } catch (error) {
            console.error('Content script error:', error);
            sendResponse({ success: false, error: error.message });
        }
    }

    addIndicator() {
        // Add visual indicator that extension is active
        const indicator = document.createElement('div');
        indicator.id = 'eta-exporter-indicator';
        indicator.innerHTML = `
            <div style="
                position: fixed;
                top: 10px;
                right: 10px;
                background: #10b981;
                color: white;
                padding: 8px 12px;
                border-radius: 6px;
                font-size: 12px;
                font-weight: bold;
                z-index: 10000;
                box-shadow: 0 2px 8px rgba(0,0,0,0.2);
                font-family: 'Segoe UI', sans-serif;
            ">
                ✓ ETA Exporter Ready
            </div>
        `;
        document.body.appendChild(indicator);

        // Remove indicator after 3 seconds
        setTimeout(() => {
            const elem = document.getElementById('eta-exporter-indicator');
            if (elem) elem.remove();
        }, 3000);
    }

    async extractInvoiceData() {
        try {
            console.log('Starting invoice data extraction...');
            
            // Check if we're on the documents page
            if (!window.location.href.includes('/documents')) {
                throw new Error('يرجى الانتقال إلى صفحة المستندات أولاً');
            }

            // Wait for page to load completely
            await this.waitForPageLoad();

            // Try multiple selectors for the table
            const tableSelectors = [
                'table',
                '.ms-DetailsList table',
                '[role="grid"]',
                '.ms-DetailsList',
                '.grid'
            ];

            let table = null;
            for (const selector of tableSelectors) {
                table = document.querySelector(selector);
                if (table) {
                    console.log(`Found table with selector: ${selector}`);
                    break;
                }
            }

            if (!table) {
                // Try to find any element that might contain invoice data
                const detailsList = document.querySelector('.ms-DetailsList');
                const listCells = document.querySelectorAll('.ms-List-cell');
                const detailsRows = document.querySelectorAll('.ms-DetailsRow');
                
                console.log('Table search results:', {
                    detailsList: !!detailsList,
                    listCells: listCells.length,
                    detailsRows: detailsRows.length
                });

                if (detailsRows.length > 0) {
                    console.log('Using DetailsRow elements instead of table');
                    const invoices = this.parseDetailsRows(detailsRows);
                    const pagination = this.extractPaginationInfo();
                    
                    return {
                        success: true,
                        data: {
                            invoices: invoices,
                            totalCount: pagination.totalCount,
                            currentPage: pagination.currentPage,
                            totalPages: pagination.totalPages
                        }
                    };
                }

                throw new Error('لم يتم العثور على جدول البيانات. تأكد من تحميل الصفحة بالكامل.');
            }

            const invoices = this.parseInvoiceTable(table);
            const pagination = this.extractPaginationInfo();

            console.log(`Extracted ${invoices.length} invoices`);

            return {
                success: true,
                data: {
                    invoices: invoices,
                    totalCount: pagination.totalCount,
                    currentPage: pagination.currentPage,
                    totalPages: pagination.totalPages
                }
            };
        } catch (error) {
            console.error('Error in extractInvoiceData:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    async waitForPageLoad() {
        // Wait for the page to be fully loaded
        if (document.readyState !== 'complete') {
            await new Promise(resolve => {
                window.addEventListener('load', resolve);
            });
        }

        // Wait for React/Angular to render
        await new Promise(resolve => setTimeout(resolve, 2000));

        // Wait for any loading indicators to disappear
        const loadingSelectors = [
            '.LoadingIndicator',
            '.ms-Spinner',
            '.loading',
            '[data-is-loading="true"]'
        ];

        for (const selector of loadingSelectors) {
            const loadingElement = document.querySelector(selector);
            if (loadingElement && loadingElement.style.display !== 'none') {
                await this.waitForElementToDisappear(selector, 10000);
            }
        }
    }

    parseDetailsRows(detailsRows) {
        const invoices = [];
        
        detailsRows.forEach((row, index) => {
            try {
                const cells = row.querySelectorAll('.ms-DetailsRow-cell, [role="rowheader"], [role="gridcell"]');
                if (cells.length === 0) return;

                // Extract data from the row structure
                const invoice = this.extractInvoiceFromRow(row, cells, index);
                if (invoice) {
                    invoices.push(invoice);
                }
            } catch (error) {
                console.warn(`Error parsing details row ${index}:`, error);
            }
        });

        return invoices;
    }

    extractInvoiceFromRow(row, cells, index) {
        try {
            // Look for specific data in the row
            const invoice = {
                serialNumber: index + 1,
                detailsButton: 'عرض',
                documentType: 'فاتورة',
                documentVersion: '1.0',
                status: '',
                issueDate: '',
                submissionDate: '',
                invoiceCurrency: 'EGP',
                invoiceValue: '',
                vatAmount: '',
                taxDiscount: '0',
                totalAmount: '',
                internalNumber: '',
                electronicNumber: '',
                sellerTaxNumber: '',
                sellerName: '',
                sellerAddress: 'غير محدد',
                buyerTaxNumber: '',
                buyerName: '',
                buyerAddress: 'غير محدد',
                purchaseOrderRef: '',
                purchaseOrderDesc: '',
                salesOrderRef: '',
                electronicSignature: 'موقع إلكترونياً',
                foodDrugGuide: '',
                externalLink: ''
            };

            // Try to extract data from various possible locations in the row
            const allText = row.textContent || '';
            
            // Look for patterns in the text
            const patterns = {
                electronicNumber: /([A-Z0-9]{26})/,
                internalNumber: /Internal ID:\s*(\d+)/i,
                totalAmount: /(\d+(?:,\d{3})*(?:\.\d{2})?)\s*EGP/,
                issueDate: /(\d{1,2}\/\d{1,2}\/\d{4})/
            };

            // Extract electronic number
            const electronicMatch = allText.match(patterns.electronicNumber);
            if (electronicMatch) {
                invoice.electronicNumber = electronicMatch[1];
            }

            // Extract internal number
            const internalMatch = allText.match(patterns.internalNumber);
            if (internalMatch) {
                invoice.internalNumber = internalMatch[1];
            }

            // Extract total amount
            const amountMatch = allText.match(patterns.totalAmount);
            if (amountMatch) {
                invoice.totalAmount = amountMatch[1];
                invoice.invoiceValue = amountMatch[1];
            }

            // Extract date
            const dateMatch = allText.match(patterns.issueDate);
            if (dateMatch) {
                invoice.issueDate = dateMatch[1];
                invoice.submissionDate = dateMatch[1];
            }

            // Try to extract from specific cell structures
            cells.forEach((cell, cellIndex) => {
                const cellText = cell.textContent?.trim() || '';
                
                // Look for specific data patterns in cells
                if (cellText.includes('Valid') || cellText.includes('صالح')) {
                    invoice.status = 'Valid';
                } else if (cellText.includes('Invalid') || cellText.includes('غير صالح')) {
                    invoice.status = 'Invalid';
                }

                // Look for seller/buyer names
                if (cellText.length > 10 && !cellText.match(/^\d+/) && !cellText.includes('EGP')) {
                    if (cellIndex < cells.length / 2) {
                        invoice.sellerName = invoice.sellerName || cellText;
                    } else {
                        invoice.buyerName = invoice.buyerName || cellText;
                    }
                }
            });

            return invoice;
        } catch (error) {
            console.warn('Error extracting invoice from row:', error);
            return null;
        }
    }

    parseInvoiceTable(table) {
        const invoices = [];
        const rows = table.querySelectorAll('tbody tr, .ms-List-cell, .ms-DetailsRow');

        rows.forEach((row, index) => {
            try {
                const cells = row.querySelectorAll('td, .ms-DetailsRow-cell, [role="gridcell"]');
                if (cells.length === 0) return;

                const invoice = {
                    serialNumber: index + 1,
                    detailsButton: this.extractDetailsButton(cells[1]),
                    documentType: this.extractCellText(cells[2]) || 'فاتورة',
                    documentVersion: this.extractCellText(cells[3]) || '1.0',
                    status: this.extractCellText(cells[4]),
                    issueDate: this.extractCellText(cells[5]),
                    submissionDate: this.extractCellText(cells[6]),
                    invoiceCurrency: this.extractCellText(cells[7]) || 'EGP',
                    invoiceValue: this.extractCellText(cells[8]),
                    vatAmount: this.extractCellText(cells[9]),
                    taxDiscount: this.extractCellText(cells[10]) || '0',
                    totalAmount: this.extractCellText(cells[11]),
                    internalNumber: this.extractCellText(cells[12]),
                    electronicNumber: this.extractCellText(cells[13]),
                    sellerTaxNumber: this.extractCellText(cells[14]),
                    sellerName: this.extractCellText(cells[15]),
                    sellerAddress: 'غير محدد',
                    buyerTaxNumber: this.extractCellText(cells[16]),
                    buyerName: this.extractCellText(cells[17]),
                    buyerAddress: 'غير محدد',
                    purchaseOrderRef: this.extractCellText(cells[18]),
                    purchaseOrderDesc: this.extractCellText(cells[19]),
                    salesOrderRef: this.extractCellText(cells[20]),
                    electronicSignature: this.extractCellText(cells[21]) || 'موقع إلكترونياً',
                    foodDrugGuide: this.extractCellText(cells[22]),
                    externalLink: this.extractCellText(cells[23])
                };

                invoices.push(invoice);
            } catch (error) {
                console.warn(`Error parsing row ${index}:`, error);
            }
        });

        return invoices;
    }

    extractCellText(cell) {
        if (!cell) return '';
        
        // Try to get text from various possible structures
        const textContent = cell.textContent || cell.innerText || '';
        return textContent.trim();
    }

    extractDetailsButton(cell) {
        if (!cell) return '';
        
        const button = cell.querySelector('button, a');
        if (button) {
            return button.textContent?.trim() || 'عرض';
        }
        
        return this.extractCellText(cell) || 'عرض';
    }

    extractPaginationInfo() {
        let totalCount = 0;
        let currentPage = 1;
        let totalPages = 1;

        try {
            // Look for pagination elements
            const paginationSelectors = [
                '.pagination-info',
                '.page-info',
                '.ms-Paging',
                '[data-automation-id="paging"]'
            ];

            for (const selector of paginationSelectors) {
                const paginationElement = document.querySelector(selector);
                if (paginationElement) {
                    const text = paginationElement.textContent;
                    const matches = text.match(/(\d+)/g);
                    if (matches && matches.length >= 3) {
                        currentPage = parseInt(matches[0]) || 1;
                        totalPages = parseInt(matches[1]) || 1;
                        totalCount = parseInt(matches[2]) || 0;
                        break;
                    }
                }
            }

            // Alternative: count visible rows
            const visibleRows = document.querySelectorAll('.ms-DetailsRow, tbody tr');
            if (visibleRows.length > 0) {
                totalCount = Math.max(totalCount, visibleRows.length);
            }

            // Look for any counter text
            const allElements = document.querySelectorAll('*');
            for (const element of allElements) {
                const text = element.textContent;
                if (text && (text.includes('total') || text.includes('إجمالي') || text.includes('Total'))) {
                    const numbers = text.match(/\d+/g);
                    if (numbers) {
                        const largestNumber = Math.max(...numbers.map(n => parseInt(n)));
                        if (largestNumber > totalCount && largestNumber < 100000) {
                            totalCount = largestNumber;
                        }
                    }
                }
            }

        } catch (error) {
            console.warn('Error extracting pagination info:', error);
        }

        return {
            totalCount: totalCount || 50, // Default fallback
            currentPage,
            totalPages: Math.max(totalPages, 1)
        };
    }

    async extractAddressesForInvoice(invoiceId, options) {
        try {
            console.log(`Extracting addresses for invoice: ${invoiceId}`);
            
            // Store current URL to return to later
            const originalUrl = window.location.href;
            
            // Navigate to invoice detail page
            const detailUrl = `/documents/${invoiceId}`;
            
            return new Promise((resolve, reject) => {
                const timeout = setTimeout(() => {
                    window.location.href = originalUrl;
                    resolve({
                        success: false,
                        addresses: {
                            sellerAddress: 'غير محدد',
                            buyerAddress: 'غير محدد'
                        }
                    });
                }, 30000);

                // Navigate to detail page
                window.location.href = detailUrl;
                
                // Wait for page to load and extract addresses
                const checkForAddresses = async () => {
                    try {
                        // Wait for the detail page to load
                        await this.waitForElement('.issuer, .receiver, textarea', 5000);
                        
                        const addresses = {
                            sellerAddress: 'غير محدد',
                            buyerAddress: 'غير محدد'
                        };

                        // Extract seller address with multiple selectors
                        if (options.sellerAddress) {
                            const sellerSelectors = [
                                '#TextField49',
                                'textarea[id*="TextField49"]',
                                '.issuer textarea',
                                '.issuer .ms-TextField-field',
                                'textarea[rows="4"]'
                            ];

                            for (const selector of sellerSelectors) {
                                const sellerField = document.querySelector(selector);
                                if (sellerField && sellerField.value) {
                                    addresses.sellerAddress = this.cleanAddressText(sellerField.value);
                                    break;
                                }
                            }
                        }

                        // Extract buyer address with multiple selectors
                        if (options.buyerAddress) {
                            const buyerSelectors = [
                                '#TextField64',
                                'textarea[id*="TextField64"]',
                                '.receiver textarea',
                                '.receiverAddress textarea',
                                '.receiver .ms-TextField-field'
                            ];

                            for (const selector of buyerSelectors) {
                                const buyerField = document.querySelector(selector);
                                if (buyerField && buyerField.value) {
                                    addresses.buyerAddress = this.cleanAddressText(buyerField.value);
                                    break;
                                }
                            }
                        }

                        clearTimeout(timeout);
                        
                        // Navigate back to original page
                        setTimeout(() => {
                            window.location.href = originalUrl;
                            resolve({
                                success: true,
                                addresses: addresses
                            });
                        }, 1000);

                    } catch (error) {
                        clearTimeout(timeout);
                        window.location.href = originalUrl;
                        resolve({
                            success: false,
                            addresses: {
                                sellerAddress: 'غير محدد',
                                buyerAddress: 'غير محدد'
                            }
                        });
                    }
                };

                // Start checking after a delay to allow navigation
                setTimeout(checkForAddresses, 3000);
            });

        } catch (error) {
            console.error('Error extracting addresses:', error);
            return {
                success: false,
                addresses: {
                    sellerAddress: 'غير محدد',
                    buyerAddress: 'غير محدد'
                }
            };
        }
    }

    cleanAddressText(address) {
        if (!address || typeof address !== 'string') {
            return 'غير محدد';
        }
        
        const cleaned = address
            .trim()
            .replace(/\s+/g, ' ')
            .replace(/\n+/g, '\n')
            .replace(/^\n|\n$/g, '');
        
        return cleaned || 'غير محدد';
    }

    async extractAllPagesData(options = {}) {
        try {
            const allInvoices = [];
            let currentPage = 1;
            const totalPages = this.extractPaginationInfo().totalPages;

            // For now, just return current page data
            // Full multi-page extraction would require more complex navigation
            const currentPageData = await this.extractInvoiceData();
            
            if (currentPageData.success) {
                return {
                    success: true,
                    data: currentPageData.data.invoices
                };
            } else {
                throw new Error(currentPageData.error);
            }
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    async extractInvoiceDetails(invoiceId) {
        try {
            // This would navigate to the invoice detail page and extract line items
            // For now, return a placeholder structure
            return {
                success: true,
                data: [
                    {
                        itemCode: invoiceId,
                        description: 'إجمالي قيمة الفاتورة',
                        unitCode: 'EA',
                        unitName: 'فاتورة',
                        quantity: '1',
                        unitPrice: '0',
                        totalValue: '0',
                        taxAmount: '0',
                        vatAmount: '0',
                        totalWithVat: '0'
                    }
                ]
            };
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    async waitForElement(selector, timeout = 5000) {
        return new Promise((resolve, reject) => {
            const element = document.querySelector(selector);
            if (element) {
                resolve(element);
                return;
            }

            const observer = new MutationObserver((mutations, obs) => {
                const element = document.querySelector(selector);
                if (element) {
                    obs.disconnect();
                    resolve(element);
                }
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true
            });

            setTimeout(() => {
                observer.disconnect();
                reject(new Error(`Element ${selector} not found within ${timeout}ms`));
            }, timeout);
        });
    }

    async waitForElementToDisappear(selector, timeout = 5000) {
        return new Promise((resolve) => {
            const checkElement = () => {
                const element = document.querySelector(selector);
                if (!element || element.style.display === 'none' || element.style.visibility === 'hidden') {
                    resolve();
                    return;
                }
                setTimeout(checkElement, 100);
            };
            
            checkElement();
            
            // Timeout fallback
            setTimeout(resolve, timeout);
        });
    }

    showLoadingOverlay(message) {
        const overlay = document.createElement('div');
        overlay.className = 'eta-loading-overlay';
        overlay.innerHTML = `
            <div style="text-align: center; color: #374151;">
                <div class="eta-loading-spinner"></div>
                <div style="margin-top: 16px; font-size: 16px; font-weight: 500;">
                    ${message}
                </div>
            </div>
        `;
        document.body.appendChild(overlay);
        return overlay;
    }

    hideLoadingOverlay(overlay) {
        if (overlay && overlay.parentNode) {
            overlay.parentNode.removeChild(overlay);
        }
    }
}

// Initialize content script
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        new ETAContentScript();
    });
} else {
    new ETAContentScript();
}

// Handle page navigation
let lastUrl = location.href;
new MutationObserver(() => {
    const url = location.href;
    if (url !== lastUrl) {
        lastUrl = url;
        // Reinitialize on navigation
        setTimeout(() => {
            new ETAContentScript();
        }, 1000);
    }
}).observe(document, { subtree: true, childList: true });