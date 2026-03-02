/**
 * Azure Speech Service Connection Test
 * Verifies that the speech service credentials are valid and the service is accessible
 */

import * as https from 'https';

// Load environment variables from .env.local.user for testing
import * as dotenv from 'dotenv';
import * as path from 'path';

dotenv.config({ path: path.join(__dirname, '..', 'env', '.env.local.user') });

interface TestResult {
  success: boolean;
  message: string;
  details?: any;
}

/**
 * Test Azure Speech Service token endpoint
 * This verifies the subscription key is valid by attempting to get an access token
 */
async function testSpeechServiceConnection(): Promise<TestResult> {
  const speechKey = process.env.SECRET_AZURE_SPEECH_KEY || process.env.AZURE_SPEECH_KEY;
  const speechRegion = process.env.AZURE_SPEECH_REGION;

  if (!speechKey) {
    return {
      success: false,
      message: 'Speech service key not found in environment variables',
      details: 'Expected SECRET_AZURE_SPEECH_KEY or AZURE_SPEECH_KEY'
    };
  }

  if (!speechRegion) {
    return {
      success: false,
      message: 'Speech service region not found in environment variables',
      details: 'Expected AZURE_SPEECH_REGION'
    };
  }

  console.log(`Testing connection to Azure Speech Service in region: ${speechRegion}`);
  console.log(`API Key: ${speechKey.substring(0, 8)}...`);

  return new Promise((resolve) => {
    const tokenEndpoint = `https://${speechRegion}.api.cognitive.microsoft.com/sts/v1.0/issueToken`;
    
    const options = {
      method: 'POST',
      headers: {
        'Ocp-Apim-Subscription-Key': speechKey,
        'Content-Length': 0
      }
    };

    const req = https.request(tokenEndpoint, options, (res) => {
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {
        if (res.statusCode === 200) {
          resolve({
            success: true,
            message: 'Azure Speech Service connection successful! ✓',
            details: {
              region: speechRegion,
              endpoint: tokenEndpoint,
              statusCode: res.statusCode,
              tokenReceived: data.length > 0
            }
          });
        } else {
          resolve({
            success: false,
            message: `Azure Speech Service returned error status ${res.statusCode}`,
            details: {
              statusCode: res.statusCode,
              statusMessage: res.statusMessage,
              response: data,
              endpoint: tokenEndpoint
            }
          });
        }
      });
    });

    req.on('error', (error) => {
      resolve({
        success: false,
        message: 'Failed to connect to Azure Speech Service',
        details: {
          error: error.message,
          endpoint: tokenEndpoint
        }
      });
    });

    req.end();
  });
}

/**
 * Run the test
 */
async function runTest() {
  console.log('='.repeat(60));
  console.log('Azure Speech Service Connection Test');
  console.log('='.repeat(60));
  console.log();

  const result = await testSpeechServiceConnection();

  console.log();
  if (result.success) {
    console.log('✅ TEST PASSED');
    console.log(result.message);
  } else {
    console.log('❌ TEST FAILED');
    console.log(result.message);
  }

  if (result.details) {
    console.log();
    console.log('Details:');
    console.log(JSON.stringify(result.details, null, 2));
  }

  console.log();
  console.log('='.repeat(60));

  process.exit(result.success ? 0 : 1);
}

// Run if executed directly
if (require.main === module) {
  runTest().catch((error) => {
    console.error('Test execution failed:', error);
    process.exit(1);
  });
}

export { testSpeechServiceConnection };
