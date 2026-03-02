/**
 * Config Loader Test
 * Tests the config module to ensure all environment variables load correctly
 */

import * as dotenv from 'dotenv';
import * as path from 'path';

// Load environment variables
dotenv.config({ path: path.join(__dirname, '..', 'env', '.env.local.user') });

// Import config after loading env vars
import { loadConfig, clearConfigCache } from '../src/config';

async function testConfigLoader() {
  console.log('='.repeat(60));
  console.log('Configuration Loader Test');
  console.log('='.repeat(60));
  console.log();

  try {
    // Clear cache to force fresh load
    clearConfigCache();
    
    // Load configuration
    const config = loadConfig();

    console.log();
    console.log('✅ Configuration loaded successfully!');
    console.log();
    console.log('Loaded configuration summary:');
    console.log('  - Azure Tenant ID:', config.azureTenantId);
    console.log('  - Azure Client ID:', config.azureClientId);
    console.log('  - Bot App ID:', config.botAppId || '(empty - will be populated during debug)');
    console.log('  - Graph Base URL:', config.graphBaseUrl);
    console.log('  - Azure OpenAI Endpoint:', config.azureOpenAIEndpoint);
    console.log('  - Azure OpenAI Deployment:', config.azureOpenAIDeployment);
    console.log('  - Azure Speech Region:', config.azureSpeechRegion);
    console.log('  - SQL Server:', config.sqlServer);
    console.log('  - SQL Database:', config.sqlDatabase);
    console.log();
    console.log('='.repeat(60));

    return true;
  } catch (error) {
    console.log();
    console.log('❌ Configuration load failed!');
    console.log();
    if (error instanceof Error) {
      console.log(error.message);
    } else {
      console.log(error);
    }
    console.log();
    console.log('='.repeat(60));

    return false;
  }
}

// Run test
testConfigLoader().then((success) => {
  process.exit(success ? 0 : 1);
});
