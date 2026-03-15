/**
 * 🌐 API Manager - Handles all API communications and integrations
 */

class APIManager {
    constructor() {
        this.endpoints = new Map();
        this.cache = new Map();
        this.requestQueue = [];
        this.isOnline = navigator.onLine;
        this.retryAttempts = 3;
        this.timeout = 30000; // 30 seconds
        
        this.setupNetworkListeners();
    }

    init() {
        this.loadEndpoints();
        this.setupInterceptors();
    }

    setupNetworkListeners() {
        window.addEventListener('online', () => {
            this.isOnline = true;
            this.processRequestQueue();
        });

        window.addEventListener('offline', () => {
            this.isOnline = false;
        });
    }

    loadEndpoints() {
        const config = window.app?.config?.services;
        if (config) {
            Object.entries(config).forEach(([key, service]) => {
                if (service.url) {
                    this.endpoints.set(key, service.url);
                }
            });
        }
    }

    setupInterceptors() {
        // Add global request/response interceptors if needed
        const originalFetch = window.fetch;
        window.fetch = async (...args) => {
            const [url, options = {}] = args;
            
            // Add common headers
            const headers = {
                'Content-Type': 'application/json',
                'X-Dashboard-Version': '1.0.0',
                ...options.headers
            };
            
            const config = {
                ...options,
                headers,
                timeout: this.timeout
            };
            
            // Add timeout handling
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), this.timeout);
            config.signal = controller.signal;
            
            try {
                const response = await originalFetch(url, config);
                clearTimeout(timeoutId);
                
                // Global response handling
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }
                
                return response;
            } catch (error) {
                clearTimeout(timeoutId);
                
                if (error.name === 'AbortError') {
                    throw new Error('Request timeout');
                }
                
                throw error;
            }
        };
    }

    // Google Apps Script Integration
    // NOTE: This calls the backend Google Apps Script via doPost()
    // The script must be deployed as a web app with:
    // - Execute as: User accessing the web app
    // - Who has access: Anyone (or specific users)
    async callGoogleScript(scriptId, functionName, parameters = []) {
        // Use the inventory endpoint as the primary backend
        // All functions are routed through the same doPost() endpoint
        const endpoint = this.endpoints.get('inventory');

        if (!endpoint) {
            throw new Error('No Google Apps Script endpoint configured. Please add your script URL in config.json under services.inventory.url');
        }

        const requestData = {
            function: functionName,
            parameters: parameters,
            devMode: false
        };

        try {
            const response = await this.makeRequest('POST', endpoint, requestData);
            return this.handleGoogleScriptResponse(response);
        } catch (error) {
            console.error('Google Apps Script call failed:', error);
            throw error;
        }
    }

    async handleGoogleScriptResponse(response) {
        const data = await response.json();

        if (!data.success || data.error) {
            throw new Error(`Google Apps Script Error: ${data.error?.message || 'Unknown error'}`);
        }

        return data.response;
    }

    // Generic HTTP methods
    async makeRequest(method, url, data = null, options = {}) {
        const cacheKey = `${method}:${url}:${JSON.stringify(data)}`;
        
        // Check cache for GET requests
        if (method === 'GET' && this.cache.has(cacheKey)) {
            const cached = this.cache.get(cacheKey);
            if (Date.now() - cached.timestamp < 300000) { // 5 minutes
                return Promise.resolve(cached.response);
            }
        }
        
        const request = {
            method,
            url,
            data,
            options,
            cacheKey,
            attempts: 0,
            maxAttempts: this.retryAttempts
        };
        
        if (!this.isOnline) {
            // Queue request for when online
            this.requestQueue.push(request);
            throw new Error('No internet connection. Request queued for retry.');
        }
        
        return this.executeRequest(request);
    }

    async executeRequest(request) {
        const { method, url, data, options, cacheKey } = request;
        
        const fetchOptions = {
            method,
            ...options
        };
        
        if (data && method !== 'GET') {
            fetchOptions.body = JSON.stringify(data);
        }
        
        try {
            const response = await fetch(url, fetchOptions);
            
            // Cache successful GET responses
            if (method === 'GET' && response.ok) {
                this.cache.set(cacheKey, {
                    response: response.clone(),
                    timestamp: Date.now()
                });
            }
            
            return response;
        } catch (error) {
            return this.handleRequestError(request, error);
        }
    }

    async handleRequestError(request, error) {
        request.attempts++;
        
        // Retry logic
        if (request.attempts < request.maxAttempts && this.shouldRetry(error)) {
            console.warn(`Request failed, retrying (${request.attempts}/${request.maxAttempts}):`, error.message);
            
            // Exponential backoff
            const delay = Math.pow(2, request.attempts - 1) * 1000;
            await this.sleep(delay);
            
            return this.executeRequest(request);
        }
        
        // All retries exhausted
        console.error('Request failed after all retries:', error);
        throw error;
    }

    shouldRetry(error) {
        // Retry on network errors, timeouts, and server errors (5xx)
        return (
            error.name === 'TypeError' || // Network error
            error.message.includes('timeout') ||
            error.message.includes('fetch')
        );
    }

    async processRequestQueue() {
        if (this.requestQueue.length === 0) return;
        
        console.log(`Processing ${this.requestQueue.length} queued requests...`);
        
        const requests = [...this.requestQueue];
        this.requestQueue = [];
        
        for (const request of requests) {
            try {
                await this.executeRequest(request);
                console.log('Queued request completed:', request.url);
            } catch (error) {
                console.error('Queued request failed:', error);
                // Could re-queue or notify user
            }
        }
    }

    // Tool-specific API methods
    async searchInventory(query) {
        try {
            return await this.callGoogleScript('inventory', 'askInventory', [query]);
        } catch (error) {
            console.error('Inventory search failed:', error);
            return {
                answer: 'Search temporarily unavailable. Please try again later.',
                source: 'error',
                success: false
            };
        }
    }

    async browseInventory() {
        try {
            return await this.callGoogleScript('inventory', 'browseInventory', []);
        } catch (error) {
            console.error('Browse inventory failed:', error);
            return { items: [], total: 0 };
        }
    }

    async browseInventoryPaginated(params) {
        try {
            return await this.callGoogleScript('inventory', 'browseInventoryPaginated', [params]);
        } catch (error) {
            console.error('Paginated browse failed:', error);
            return { items: [], total: 0, page: 1, pageSize: 50, totalPages: 0 };
        }
    }

    async updateInventory(updateData) {
        try {
            return await this.callGoogleScript('inventory', 'updateInventory', [updateData]);
        } catch (error) {
            console.error('Inventory update failed:', error);
            throw error;
        }
    }

    async gradeProduct(productData) {
        try {
            return await this.callGoogleScript('grading', 'gradeProduct', [productData]);
        } catch (error) {
            console.error('Product grading failed:', error);
            throw error;
        }
    }

    async getSchedule(date) {
        try {
            return await this.callGoogleScript('scheduler', 'getSchedule', [date]);
        } catch (error) {
            console.error('Schedule fetch failed:', error);
            throw error;
        }
    }

    async updateSchedule(scheduleData) {
        try {
            return await this.callGoogleScript('scheduler', 'updateSchedule', [scheduleData]);
        } catch (error) {
            console.error('Schedule update failed:', error);
            throw error;
        }
    }

    async checkoutTool(toolData) {
        try {
            return await this.callGoogleScript('tools', 'checkoutTool', [toolData]);
        } catch (error) {
            console.error('Tool checkout failed:', error);
            throw error;
        }
    }

    async returnTool(toolData) {
        try {
            return await this.callGoogleScript('tools', 'returnTool', [toolData]);
        } catch (error) {
            console.error('Tool return failed:', error);
            throw error;
        }
    }

    // User authentication
    async getUserInfo() {
        try {
            return await this.callGoogleScript('auth', 'getUserInfo', []);
        } catch (error) {
            console.warn('Could not get user info:', error);
            return {
                name: 'Guest User',
                email: 'guest@deeproots.com',
                avatar: '👤'
            };
        }
    }

    async checkAccess() {
        try {
            return await this.callGoogleScript('auth', 'checkUserAccess', []);
        } catch (error) {
            console.warn('Access check failed:', error);
            return { hasAccess: true, role: 'guest' };
        }
    }

    // Data export and backup
    async exportData(type) {
        try {
            const functionName = `export${type.charAt(0).toUpperCase() + type.slice(1)}CSV`;
            return await this.callGoogleScript('inventory', functionName, []);
        } catch (error) {
            console.error('Data export failed:', error);
            throw error;
        }
    }

    async createBackup() {
        try {
            return await this.callGoogleScript('inventory', 'createDataBackup', []);
        } catch (error) {
            console.error('Backup creation failed:', error);
            throw error;
        }
    }

    async generateReport() {
        try {
            return await this.callGoogleScript('inventory', 'generateComprehensiveReport', []);
        } catch (error) {
            console.error('Report generation failed:', error);
            throw error;
        }
    }

    // WebRTC for real-time features (if needed)
    async establishRealTimeConnection() {
        // Placeholder for WebRTC or WebSocket connections
        // Could be used for real-time inventory updates, notifications, etc.
        console.log('Real-time connection placeholder');
    }

    // Utility methods
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    clearCache() {
        this.cache.clear();
        console.log('API cache cleared');
    }

    getCacheInfo() {
        return {
            size: this.cache.size,
            entries: Array.from(this.cache.keys())
        };
    }

    // Health check
    async healthCheck() {
        const results = {};
        
        for (const [name, endpoint] of this.endpoints.entries()) {
            try {
                const start = Date.now();
                await fetch(endpoint, { method: 'HEAD', timeout: 5000 });
                results[name] = {
                    status: 'healthy',
                    responseTime: Date.now() - start
                };
            } catch (error) {
                results[name] = {
                    status: 'unhealthy',
                    error: error.message
                };
            }
        }
        
        return results;
    }

    // Request statistics
    getStats() {
        return {
            cacheSize: this.cache.size,
            queueLength: this.requestQueue.length,
            isOnline: this.isOnline,
            endpoints: Array.from(this.endpoints.keys())
        };
    }
}

// Export for module systems
if (typeof module !== 'undefined' && module.exports) {
    module.exports = APIManager;
}