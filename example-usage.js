// @ts-nocheck
import logger from './logger.js';

// Contoh penggunaan dalam GSpread Appender
async function simulateGSpreadOperations() {
  console.log('🚀 Starting GSpread Logger Example...\n');

  // 1. BASIC LOGGING
  logger.info('Application started');
  logger.debug('Debug information', { config: 'loaded', version: '1.0.0' });

  // 2. FILE PROCESSING EXAMPLE
  logger.fileProcessing('/path/to/server1.txt', 'started', { size: '150KB' });
  
  // Data validation
  logger.dataValidation('Hostname', 'srv-web-01', 'valid');
  logger.dataValidation('IP Address', '192.168.1.100', 'invalid');
  
  // Performance tracking
  const timer = logger.startTimer('file_parsing');
  await new Promise(resolve => setTimeout(resolve, 500)); // Simulate work
  timer.end({ linesProcessed: 150, entitiesFound: 45 });

  // 3. SPREADSHEET OPERATIONS
  logger.spreadsheetOperation('update', '1ABC123SpreadsheetID', {
    range: 'Sheet1!A1:Z100',
    cellsUpdated: 45
  });

  // 4. BATCH PROCESSING
  logger.batchOperation(100, 95, 3, 2345);

  // 5. VM OPERATIONS
  const vmData = { 
    'Hostname': 'vm-database-01', 
    'Serial Number': 'VMware-42a1b2c3' 
  };
  
  const parentData = { 
    'Hostname': 'srv-hypervisor-01', 
    'Serial Number': 'SN-PHYSICAL-001' 
  };
  
  logger.vmOperation(vmData, parentData, 'appended_below_parent');

  // 6. OPERATION TRACKING (LONG RUNNING)
  const batchOpId = logger.startOperation('nightly_batch_processing', {
    filesCount: 250,
    estimatedDuration: '5 minutes'
  });

  // Simulate batch work
  await new Promise(resolve => setTimeout(resolve, 1000));
  
  // Complete operation successfully
  logger.endOperation(batchOpId, {
    processed: 245,
    failed: 5,
    totalDuration: '1.2 seconds'
  });

  // 7. ERROR HANDLING
  try {
    throw new Error('Failed to connect to Google API');
  } catch (error) {
    const errorOpId = logger.startOperation('api_connection');
    logger.errorOperation(errorOpId, error, {
      retryCount: 3,
      endpoint: 'https://sheets.googleapis.com/v4/spreadsheets'
    });
  }

  // 8. COMPLEX SCENARIO - MULTIPLE OPERATIONS
  const mainOpId = logger.startOperation('main_data_sync');
  
  const subOps = [
    logger.startOperation('user_data_extraction'),
    logger.startOperation('server_inventory'),
    logger.startOperation('network_mapping')
  ];

  // Simulate parallel operations
  await Promise.all(subOps.map(async (opId, index) => {
    await new Promise(resolve => setTimeout(resolve, 300 * (index + 1)));
    logger.endOperation(opId, { records: (index + 1) * 50 });
  }));

  // Complete main operation
  logger.endOperation(mainOpId, {
    totalRecords: 300,
    status: 'completed_successfully'
  });

  // 9. SHOW STATISTICS
  const stats = logger.getStats();
  console.log('\n📊 Active Operations:', stats.activeOperations);

  // 10. CLEANUP OLD LOGS (biasanya dijalankan periodic)
  await logger.cleanOldLogs();

  logger.info('All operations completed successfully');
}

// Run the example
simulateGSpreadOperations().catch(error => {
  logger.error('Example failed', { error: error.message });
});

export { simulateGSpreadOperations };