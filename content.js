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
            // Check if we're on the documents page
            if (!window.location.href.includes('/documents')) {
                throw new Error('يرجى الانتقال إلى صفحة المستندات أولاً');
            }

            // Wait for table to load
            await this.waitForElement('table', 10000);

            const table = document.querySelector('table');
            if (!table) {
                throw new Error('لم يتم العثور على جدول البيانات');
            }

            const invoices = this.parseInvoiceTable(table);
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
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    parseInvoiceTable(table) {
        const invoices = [];
        const rows = table.querySelectorAll('tbody tr');

        rows.forEach((row, index) => {
            try {
                const cells = row.querySelectorAll('td');
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
                    sellerAddress: 'غير محدد', // Default value, will be extracted separately if needed
                    buyerTaxNumber: this.extractCellText(cells[16]),
                    buyerName: this.extractCellText(cells[17]),
                    buyerAddress: 'غير محدد', // Default value, will be extracted separately if needed
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
        // Try to extract pagination information from the page
        let totalCount = 0;
        let currentPage = 1;
        let totalPages = 1;

        try {
            // Look for pagination elements (this may vary based on the actual page structure)
            const paginationText = document.querySelector('.pagination-info, .page-info');
            if (paginationText) {
                const text = paginationText.textContent;
                const matches = text.match(/(\d+)/g);
                if (matches && matches.length >= 3) {
                    currentPage = parseInt(matches[0]) || 1;
                    totalPages = parseInt(matches[1]) || 1;
                    totalCount = parseInt(matches[2]) || 0;
                }
            }

            // Alternative: count table rows and estimate
            const table = document.querySelector('table');
            if (table) {
                const rows = table.querySelectorAll('tbody tr');
                totalCount = Math.max(totalCount, rows.length);
            }

            // Try to get total count from any displayed counter
            const counterElements = document.querySelectorAll('*');
            for (const element of counterElements) {
                const text = element.textContent;
                if (text && text.includes('total') || text.includes('إجمالي')) {
                    const numbers = text.match(/\d+/g);
                    if (numbers) {
                        totalCount = Math.max(totalCount, parseInt(numbers[numbers.length - 1]));
                    }
                }
            }

        } catch (error) {
            console.warn('Error extracting pagination info:', error);
        }

        return {
            totalCount: totalCount || 299, // Default fallback
            currentPage,
            totalPages
        };
    }

    async extractAddressesForInvoice(invoiceId, options) {
        try {
            console.log(`Extracting addresses for invoice: ${invoiceId}`);
            
            // Store current URL to return to later
            const originalUrl = window.location.href;
            
            // Navigate to invoice detail page
            const detailUrl = `/documents/${invoiceId}`;
            
            // Create a promise to handle navigation and data extraction
            return new Promise((resolve, reject) => {
                const timeout = setTimeout(() => {
                    reject(new Error('Timeout waiting for invoice detail page'));
                }, 30000);

                // Navigate to detail page
                window.location.href = detailUrl;
                
                // Wait for page to load and extract addresses
                const checkForAddresses = async () => {
                    try {
                        // Wait for the detail page to load
                        await this.waitForElement('.issuer, .receiver', 5000);
                        
                        const addresses = {
                            sellerAddress: 'غير محدد',
                            buyerAddress: 'غير محدد'
                        };

                        // Extract seller address
                        if (options.sellerAddress) {
                            const sellerAddressField = document.querySelector('#TextField49, textarea[id*="TextField49"], .issuer textarea, .issuer .ms-TextField-field');
                            if (sellerAddressField) {
                                addresses.sellerAddress = sellerAddressField.value || sellerAddressField.textContent || 'غير محدد';
                            }
                        }

                        // Extract buyer address  
                        if (options.buyerAddress) {
                            const buyerAddressField = document.querySelector('#TextField64, textarea[id*="TextField64"], .receiver textarea, .receiverAddress textarea, .receiver .ms-TextField-field');
                            if (buyerAddressField) {
                                addresses.buyerAddress = buyerAddressField.value || buyerAddressField.textContent || 'غير محدد';
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
                        // Navigate back to original page even on error
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

                // Start checking after a short delay to allow navigation
                setTimeout(checkForAddresses, 2000);
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

    async extractAllPagesData(options = {}) {
        try {
            const allInvoices = [];
            let currentPage = 1;
            const totalPages = this.extractPaginationInfo().totalPages;

            // Extract data from all pages
            for (let page = 1; page <= totalPages; page++) {
                // Send progress update
                if (options.progressCallback) {
                    chrome.runtime.sendMessage({
                        action: 'progressUpdate',
                        progress: {
                            currentPage: page,
                            totalPages: totalPages,
                            message: `جاري تحميل الصفحة ${page} من ${totalPages}...`
                        }
                    });
                }

                // Navigate to page if not current page
                if (page !== currentPage) {
                    await this.navigateToPage(page);
                    await this.waitForElement('table', 5000);
                }

                // Extract data from current page
                const table = document.querySelector('table');
                if (table) {
                    const pageInvoices = this.parseInvoiceTable(table);
                    allInvoices.push(...pageInvoices);
                }

                currentPage = page;
            }

            return {
                success: true,
                data: allInvoices
            };
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

    async navigateToPage(pageNumber) {
        // Implementation depends on the pagination structure of the site
        // This is a placeholder - you'll need to implement based on actual pagination
        const pageButton = document.querySelector(`[data-page="${pageNumber}"], .page-${pageNumber}`);
        if (pageButton) {
            pageButton.click();
            await this.waitForElement('table', 5000);
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