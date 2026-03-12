/**
 * Unit Test: Detect Hardcoded Values
 * This test scans the codebase for hardcoded strings, numbers, and configuration values
 * that should be dynamically loaded from external sources.
 */

function scanForHardcodedValues() {
    const hardcodedPatterns = [
        // Script IDs and API endpoints
        /script\.google\.com.*exec/gi,
        /AKfycb[\w-]+/gi,
        /macros\/s\/\w+\/exec/gi,
        
        // Hardcoded API keys (basic pattern)
        /['"]AIza[\w-]{30,}['"]/gi,
        
        // Database connection strings
        /mongodb:\/\/[^'"\s]+/gi,
        /mysql:\/\/[^'"\s]+/gi,
        /postgres:\/\/[^'"\s]+/gi,
        
        // Hardcoded URLs
        /https?:\/\/[^'"\s]+/gi,
        
        // Email addresses (excluding test/example emails)
        /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi,
        
        // IP addresses (excluding localhost)
        /\b(?!127\.0\.0\.1|localhost)(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b/gi,
        
        // Hardcoded file paths (Windows and Unix)
        /[C-Z]:\\[^'"\s]+/gi,
        /\/[^'"\s]+\/[\w-]+/gi,
        
        // Phone numbers
        /\b\d{3}[-.]?\d{3}[-.]?\d{4}\b/gi,
        
        // Credit card numbers (basic pattern)
        /\b\d{4}[\s-]?\d{4}[\s-]?\d{4}[\s-]?\d{4}\b/gi,
        
        // Social Security Numbers (basic pattern)
        /\b\d{3}-\d{2}-\d{4}\b/gi
    ];
    
    const whitelist = [
        // Allow localhost and common test domains
        'localhost',
        '127.0.0.1',
        'example.com',
        'test.com',
        'example.org',
        'test.org',
        
        // Allow common CDN domains
        'cdnjs.cloudflare.com',
        'cdn.jsdelivr.net',
        'fonts.googleapis.com',
        'fonts.gstatic.com',
        
        // Allow common test data
        'test@example.com',
        'user@example.com',
        'admin@example.com',
        
        // Allow placeholder values
        'placeholder',
        'lorem ipsum',
        'sample',
        'demo',
        'mock',
        
        // Allow common status messages
        'success',
        'error',
        'loading',
        'pending',
        'active',
        'inactive'
    ];
    
    function isWhitelisted(value) {
        const lowerValue = value.toLowerCase();
        return whitelist.some(item => lowerValue.includes(item.toLowerCase()));
    }
    
    function scanFile(content, filename) {
        const findings = [];
        
        hardcodedPatterns.forEach((pattern, index) => {
            const matches = content.match(pattern) || [];
            matches.forEach(match => {
                if (!isWhitelisted(match)) {
                    findings.push({
                        type: getPatternName(index),
                        value: match,
                        line: getLineNumber(content, match),
                        severity: getSeverity(index)
                    });
                }
            });
        });
        
        return findings;
    }
    
    function getPatternName(index) {
        const names = [
            'Script ID/API Endpoint',
            'API Key',
            'Database Connection',
            'Hardcoded URL',
            'Email Address',
            'IP Address',
            'File Path',
            'Phone Number',
            'Credit Card',
            'Social Security Number'
        ];
        return names[index] || 'Unknown';
    }
    
    function getSeverity(index) {
        const severities = [
            'HIGH',    // Script ID/API Endpoint
            'CRITICAL', // API Key
            'HIGH',    // Database Connection
            'MEDIUM',  // Hardcoded URL
            'MEDIUM',  // Email Address
            'LOW',     // IP Address
            'LOW',     // File Path
            'LOW',     // Phone Number
            'CRITICAL', // Credit Card
            'HIGH'     // Social Security Number
        ];
        return severities[index] || 'LOW';
    }
    
    function getLineNumber(content, match) {
        const lines = content.split('\n');
        for (let i = 0; i < lines.length; i++) {
            if (lines[i].includes(match)) {
                return i + 1;
            }
        }
        return 0;
    }
    
    return {
        scanFile,
        scanFiles: function(files) {
            const results = {};
            Object.keys(files).forEach(filename => {
                results[filename] = scanFile(files[filename], filename);
            });
            return results;
        },
        generateReport: function(findings) {
            const report = {
                timestamp: new Date().toISOString(),
                totalFiles: Object.keys(findings).length,
                totalFindings: 0,
                criticalCount: 0,
                highCount: 0,
                mediumCount: 0,
                lowCount: 0,
                findings: findings
            };
            
            Object.values(findings).forEach(fileFindings => {
                report.totalFindings += fileFindings.length;
                fileFindings.forEach(finding => {
                    switch (finding.severity) {
                        case 'CRITICAL': report.criticalCount++; break;
                        case 'HIGH': report.highCount++; break;
                        case 'MEDIUM': report.mediumCount++; break;
                        case 'LOW': report.lowCount++; break;
                    }
                });
            });
            
            return report;
        }
    };
}

// Test function to scan current codebase
function runHardcodedValueTest() {
    console.log('🔍 Starting hardcoded value detection test...');
    
    const scanner = scanForHardcodedValues();
    
    // Files to scan (you can add more files here)
    const filesToScan = {
        'script.js': '',
        'Code.js': '',
        'index.html': '',
        'nav.html': '',
        'dashboard.html': '',
        'inventory.html': '',
        'reports.html': ''
    };
    
    // Mock file reading (in real implementation, read actual files)
    const mockFiles = {
        'script.js': `
            const SCRIPT_ID = 'AKfycbx4PVz0rFjlZcJBPEHRNj4nmgL7eYJJ5pFAQx3kBE1B7Y6I4VQTJCaCxl9GxskHfT_';
            const API_KEY = 'AIzaSyBsqglgmf1O9XroTBF-qw1oIlBoHlies7A';
            const DB_URL = 'mongodb://user:pass@localhost:27017/inventory';
        `,
        'config.js': `
            export const CONFIG = {
                API_ENDPOINT: 'https://api.example.com/v1',
                WEBHOOK_URL: 'https://webhook.site/unique-id',
                ADMIN_EMAIL: 'admin@company.com'
            };
        `
    };
    
    const findings = scanner.scanFiles(mockFiles);
    const report = scanner.generateReport(findings);
    
    console.log('📊 Hardcoded Value Detection Report:');
    console.log('=====================================');
    console.log(`Total Files Scanned: ${report.totalFiles}`);
    console.log(`Total Findings: ${report.totalFindings}`);
    console.log(`Critical: ${report.criticalCount}`);
    console.log(`High: ${report.highCount}`);
    console.log(`Medium: ${report.mediumCount}`);
    console.log(`Low: ${report.lowCount}`);
    console.log('');
    
    Object.keys(report.findings).forEach(filename => {
        const fileFindings = report.findings[filename];
        if (fileFindings.length > 0) {
            console.log(`📄 ${filename}:`);
            fileFindings.forEach(finding => {
                console.log(`  - ${finding.severity}: ${finding.type} at line ${finding.line}`);
                console.log(`    Value: ${finding.value}`);
            });
            console.log('');
        }
    });
    
    // Recommendations
    console.log('💡 Recommendations:');
    console.log('==================');
    if (report.criticalCount > 0) {
        console.log('⚠️  CRITICAL: Remove all API keys and sensitive data immediately!');
    }
    if (report.highCount > 0) {
        console.log('⚠️  HIGH: Move configuration values to environment variables or secure storage');
    }
    if (report.mediumCount > 0) {
        console.log('⚠️  MEDIUM: Consider externalizing URLs and email addresses to configuration files');
    }
    if (report.lowCount > 0) {
        console.log('ℹ️  LOW: Review low-severity findings for potential externalization');
    }
    
    return report;
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { scanForHardcodedValues, runHardcodedValueTest };
}

// Run test if this file is executed directly
if (typeof window !== 'undefined') {
    // Browser environment - attach to window
    window.scanForHardcodedValues = scanForHardcodedValues;
    window.runHardcodedValueTest = runHardcodedValueTest;
} else if (typeof global !== 'undefined') {
    // Node.js environment - attach to global
    global.scanForHardcodedValues = scanForHardcodedValues;
    global.runHardcodedValueTest = runHardcodedValueTest;
}

// Example usage:
// runHardcodedValueTest();