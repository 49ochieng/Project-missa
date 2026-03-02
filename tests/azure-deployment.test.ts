/**
 * Azure Deployment Test Script
 * Tests the deployed Azure app's functionality including:
 * - Database connectivity (MSSQL)
 * - Azure OpenAI API connectivity
 * - Bot endpoint availability
 * 
 * Run with: npx ts-node tests/azure-deployment.test.ts
 */

import * as https from 'https';
import * as dotenv from 'dotenv';
import * as path from 'path';
import * as sql from 'mssql';

// Load environment variables from .env.dev.user
dotenv.config({ path: path.join(__dirname, '..', 'env', '.env.dev.user') });

interface TestResult {
  name: string;
  passed: boolean;
  message: string;
  duration: number;
}

const results: TestResult[] = [];

async function httpRequest(url: string, options: https.RequestOptions = {}): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const req = https.request(url, options, (res) => {
      let body = '';
      res.on('data', (chunk) => body += chunk);
      res.on('end', () => resolve({ status: res.statusCode || 0, body }));
    });
    req.on('error', reject);
    req.setTimeout(30000, () => {
      req.destroy();
      reject(new Error('Request timeout'));
    });
    if (options.method === 'POST' && (options as any).body) {
      req.write((options as any).body);
    }
    req.end();
  });
}

async function runTest(name: string, fn: () => Promise<{ passed: boolean; message: string }>): Promise<void> {
  const start = Date.now();
  try {
    const result = await fn();
    results.push({
      name,
      passed: result.passed,
      message: result.message,
      duration: Date.now() - start,
    });
  } catch (error) {
    results.push({
      name,
      passed: false,
      message: error instanceof Error ? error.message : String(error),
      duration: Date.now() - start,
    });
  }
}

// Test 1: Bot endpoint availability
async function testBotEndpoint(): Promise<{ passed: boolean; message: string }> {
  const botUrl = process.env.PUBLIC_BASE_URL || 'https://bot161976.azurewebsites.net';
  const response = await httpRequest(botUrl);
  
  if (response.status === 200 && response.body.includes('bots')) {
    const data = JSON.parse(response.body);
    return {
      passed: true,
      message: `Bot endpoint responding. Bot ID: ${data.id || 'unknown'}`,
    };
  }
  return { passed: false, message: `Unexpected response: ${response.status}` };
}

// Test 2: Azure OpenAI API connectivity
async function testAzureOpenAI(): Promise<{ passed: boolean; message: string }> {
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT || process.env.AOAI_ENDPOINT;
  const apiKey = process.env.SECRET_AZURE_OPENAI_API_KEY || process.env.AOAI_API_KEY;
  const model = process.env.AZURE_OPENAI_DEPLOYMENT_NAME || process.env.AOAI_MODEL || 'gpt-4.1';
  
  if (!endpoint || !apiKey) {
    return { passed: false, message: 'Missing AZURE_OPENAI_ENDPOINT or API_KEY' };
  }

  const url = `${endpoint}/openai/deployments/${model}/chat/completions?api-version=2024-02-15-preview`;
  const body = JSON.stringify({
    messages: [{ role: 'user', content: 'Say "test successful" in exactly 2 words' }],
    max_tokens: 10,
  });

  const response = await httpRequest(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey,
    },
    body,
  } as any);

  if (response.status === 200) {
    const data = JSON.parse(response.body);
    const content = data.choices?.[0]?.message?.content || '';
    return {
      passed: true,
      message: `OpenAI API working. Response: "${content.substring(0, 50)}"`,
    };
  }
  
  return { 
    passed: false, 
    message: `OpenAI API error: ${response.status} - ${response.body.substring(0, 100)}` 
  };
}

// Test 3: MSSQL Database connectivity
async function testMSSQLConnection(): Promise<{ passed: boolean; message: string }> {
  
  const server = process.env.SQL_SERVER;
  const database = process.env.SQL_DATABASE;
  const username = process.env.SQL_USERNAME;
  const password = process.env.SQL_PASSWORD;
  
  if (!server || !database || !username || !password) {
    return { passed: false, message: 'Missing SQL connection parameters' };
  }

  const config: sql.config = {
    server,
    database,
    user: username,
    password,
    options: {
      encrypt: true,
      trustServerCertificate: false,
    },
  };

  try {
    const pool = await sql.connect(config);
    const result = await pool.request().query('SELECT 1 AS test');
    await pool.close();
    
    if (result.recordset?.[0]?.test === 1) {
      return { passed: true, message: `Connected to ${database}@${server}` };
    }
    return { passed: false, message: 'Query returned unexpected result' };
  } catch (error) {
    return { 
      passed: false, 
      message: `SQL connection failed: ${error instanceof Error ? error.message : error}` 
    };
  }
}

// Test 4: Check database tables exist
async function testDatabaseTables(): Promise<{ passed: boolean; message: string }> {
  
  const config: sql.config = {
    server: process.env.SQL_SERVER!,
    database: process.env.SQL_DATABASE!,
    user: process.env.SQL_USERNAME!,
    password: process.env.SQL_PASSWORD!,
    options: { encrypt: true, trustServerCertificate: false },
  };

  try {
    const pool = await sql.connect(config);
    const result = await pool.request().query(`
      SELECT TABLE_NAME 
      FROM INFORMATION_SCHEMA.TABLES 
      WHERE TABLE_TYPE = 'BASE TABLE'
    `);
    await pool.close();
    
    const tables = result.recordset.map(r => r.TABLE_NAME);
    const expectedTables = ['conversations', 'meetings', 'meeting_participants', 'transcript_chunks', 'meeting_summaries'];
    const foundTables = expectedTables.filter(t => tables.includes(t));
    
    return {
      passed: foundTables.length >= 1,
      message: `Found tables: ${tables.join(', ')}`,
    };
  } catch (error) {
    return { passed: false, message: `Failed to query tables: ${error}` };
  }
}

// Test 5: Azure Speech Service
async function testSpeechService(): Promise<{ passed: boolean; message: string }> {
  const speechKey = process.env.SECRET_AZURE_SPEECH_KEY || process.env.AZURE_SPEECH_KEY;
  const speechRegion = process.env.AZURE_SPEECH_REGION;
  
  if (!speechKey || !speechRegion) {
    return { passed: false, message: 'Missing Speech Service credentials' };
  }

  const tokenEndpoint = `https://${speechRegion}.api.cognitive.microsoft.com/sts/v1.0/issueToken`;
  
  const response = await httpRequest(tokenEndpoint, {
    method: 'POST',
    headers: {
      'Ocp-Apim-Subscription-Key': speechKey,
      'Content-Length': '0',
    },
  } as any);
  
  if (response.status === 200 && response.body.length > 0) {
    return { passed: true, message: `Speech Service token obtained (${speechRegion})` };
  }
  
  return { passed: false, message: `Speech Service error: ${response.status}` };
}

// Test 6: Test conversation storage
async function testConversationStorage(): Promise<{ passed: boolean; message: string }> {
  
  const config: sql.config = {
    server: process.env.SQL_SERVER!,
    database: process.env.SQL_DATABASE!,
    user: process.env.SQL_USERNAME!,
    password: process.env.SQL_PASSWORD!,
    options: { encrypt: true, trustServerCertificate: false },
  };

  try {
    const pool = await sql.connect(config);
    const testConvId = `test-conv-${Date.now()}`;
    
    // Insert test message
    await pool.request()
      .input('conversation_id', sql.NVarChar, testConvId)
      .input('role', sql.NVarChar, 'user')
      .input('content', sql.NVarChar, 'Test message from deployment test')
      .input('name', sql.NVarChar, 'TestUser')
      .input('timestamp', sql.DateTime, new Date())
      .input('activity_id', sql.NVarChar, `test-activity-${Date.now()}`)
      .input('blob', sql.NVarChar, '{}')
      .query(`
        INSERT INTO conversations (conversation_id, role, content, name, timestamp, activity_id, blob)
        VALUES (@conversation_id, @role, @content, @name, @timestamp, @activity_id, @blob)
      `);
    
    // Read back
    const result = await pool.request()
      .input('conversation_id', sql.NVarChar, testConvId)
      .query('SELECT * FROM conversations WHERE conversation_id = @conversation_id');
    
    // Cleanup
    await pool.request()
      .input('conversation_id', sql.NVarChar, testConvId)
      .query('DELETE FROM conversations WHERE conversation_id = @conversation_id');
    
    await pool.close();
    
    if (result.recordset.length === 1) {
      return { passed: true, message: 'Conversation storage working (insert/read/delete)' };
    }
    return { passed: false, message: 'Insert/read verification failed' };
  } catch (error) {
    return { passed: false, message: `Storage test failed: ${error}` };
  }
}

// Test 7: Test meeting storage
async function testMeetingStorage(): Promise<{ passed: boolean; message: string }> {
  
  const config: sql.config = {
    server: process.env.SQL_SERVER!,
    database: process.env.SQL_DATABASE!,
    user: process.env.SQL_USERNAME!,
    password: process.env.SQL_PASSWORD!,
    options: { encrypt: true, trustServerCertificate: false },
  };

  try {
    const pool = await sql.connect(config);
    const testMeetingId = `test-meeting-${Date.now()}`;
    
    // Insert test meeting
    await pool.request()
      .input('meeting_id', sql.NVarChar, testMeetingId)
      .input('conversation_id', sql.NVarChar, 'test-conv')
      .input('join_url', sql.NVarChar, 'https://teams.microsoft.com/test')
      .input('title', sql.NVarChar, 'Test Meeting')
      .input('organizer_aad_id', sql.NVarChar, 'test-organizer')
      .input('status', sql.NVarChar, 'joining')
      .query(`
        INSERT INTO meetings (meeting_id, conversation_id, join_url, title, organizer_aad_id, status)
        VALUES (@meeting_id, @conversation_id, @join_url, @title, @organizer_aad_id, @status)
      `);
    
    // Read back
    const result = await pool.request()
      .input('meeting_id', sql.NVarChar, testMeetingId)
      .query('SELECT * FROM meetings WHERE meeting_id = @meeting_id');
    
    // Cleanup
    await pool.request()
      .input('meeting_id', sql.NVarChar, testMeetingId)
      .query('DELETE FROM meetings WHERE meeting_id = @meeting_id');
    
    await pool.close();
    
    if (result.recordset.length === 1 && result.recordset[0].title === 'Test Meeting') {
      return { passed: true, message: 'Meeting storage working (insert/read/delete)' };
    }
    return { passed: false, message: 'Meeting insert/read verification failed' };
  } catch (error) {
    return { passed: false, message: `Meeting storage test failed: ${error}` };
  }
}

// Main test runner
async function main() {
  console.log('═'.repeat(60));
  console.log('  Azure Deployment Test Suite');
  console.log('  Bot: ' + (process.env.PUBLIC_BASE_URL || 'https://bot161976.azurewebsites.net'));
  console.log('═'.repeat(60));
  console.log();

  await runTest('1. Bot Endpoint Availability', testBotEndpoint);
  await runTest('2. Azure OpenAI API', testAzureOpenAI);
  await runTest('3. MSSQL Database Connection', testMSSQLConnection);
  await runTest('4. Database Tables', testDatabaseTables);
  await runTest('5. Azure Speech Service', testSpeechService);
  await runTest('6. Conversation Storage CRUD', testConversationStorage);
  await runTest('7. Meeting Storage CRUD', testMeetingStorage);

  console.log();
  console.log('─'.repeat(60));
  console.log('  Results');
  console.log('─'.repeat(60));

  let passedCount = 0;
  let failedCount = 0;

  for (const result of results) {
    const status = result.passed ? '✅' : '❌';
    const statusText = result.passed ? 'PASS' : 'FAIL';
    passedCount += result.passed ? 1 : 0;
    failedCount += result.passed ? 0 : 1;
    
    console.log(`${status} ${result.name} (${result.duration}ms)`);
    console.log(`   ${statusText}: ${result.message}`);
    console.log();
  }

  console.log('═'.repeat(60));
  console.log(`  Summary: ${passedCount} passed, ${failedCount} failed`);
  console.log('═'.repeat(60));

  process.exit(failedCount > 0 ? 1 : 0);
}

main().catch(console.error);
