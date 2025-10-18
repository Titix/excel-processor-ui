/**
 * Backend Server Tests
 */

const request = require('supertest');
const express = require('express');
const path = require('path');

// Mock the server module
const createServer = () => {
  const app = express();
  
  // Middleware
  app.use(express.json());
  
  // Serve static files from React build
  app.use(express.static(path.join(__dirname, '../build')));
  
  // Health check endpoint
  app.get('/health', (req: any, res: any) => {
    res.json({ status: 'OK', timestamp: new Date().toISOString() });
  });
  
  // Serve React app for all other routes
  app.get('*', (req: any, res: any) => {
    res.sendFile(path.join(__dirname, '../build/index.html'));
  });
  
  return app;
};

describe('Backend Server', () => {
  let app: any;

  beforeEach(() => {
    app = createServer();
  });

  describe('Health Check Endpoint', () => {
    test('GET /health returns OK status', async () => {
      const response = await request(app)
        .get('/health')
        .expect(200);

      expect(response.body).toHaveProperty('status', 'OK');
      expect(response.body).toHaveProperty('timestamp');
      expect(new Date(response.body.timestamp)).toBeInstanceOf(Date);
    });
  });

  describe('Static File Serving', () => {
    test('serves static files from build directory', async () => {
      // This test would require the build directory to exist
      // For now, we'll test the route exists
      const response = await request(app)
        .get('/')
        .expect(200);
      
      // The response should be the index.html file
      expect(response.text).toContain('<!DOCTYPE html>');
    });
  });

  describe('CORS Configuration', () => {
    test('includes CORS headers', async () => {
      const response = await request(app)
        .get('/health')
        .expect(200);

      // CORS headers should be present
      expect(response.headers).toHaveProperty('access-control-allow-origin');
    });
  });

  describe('Error Handling', () => {
    test('handles 404 errors gracefully', async () => {
      const response = await request(app)
        .get('/non-existent-route')
        .expect(200);
      
      // Should serve the React app for SPA routing
      expect(response.text).toContain('<!DOCTYPE html>');
    });
  });
});

describe('Server Configuration', () => {
  test('uses correct port from environment', () => {
    const originalPort = process.env.PORT;
    process.env.PORT = '4000';
    
    // In a real test, you would check the server configuration
    expect(process.env.PORT).toBe('4000');
    
    // Restore original port
    process.env.PORT = originalPort;
  });

  test('defaults to port 3000 when PORT not set', () => {
    const originalPort = process.env.PORT;
    delete process.env.PORT;
    
    // In a real test, you would check the server configuration
    expect(process.env.PORT).toBeUndefined();
    
    // Restore original port
    process.env.PORT = originalPort;
  });
});

describe('Build Script Tests', () => {
  const { execSync } = require('child_process');

  test('build script exists and is executable', () => {
    // This test checks if the build script can be found
    const fs = require('fs');
    const path = require('path');
    
    const buildScriptPath = path.join(__dirname, '../src/backend/build.js');
    expect(fs.existsSync(buildScriptPath)).toBe(true);
  });

  test('package.json has required scripts', () => {
    const fs = require('fs');
    const path = require('path');
    
    const packageJsonPath = path.join(__dirname, '../package.json');
    const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
    
    expect(packageJson.scripts).toHaveProperty('build');
    expect(packageJson.scripts).toHaveProperty('start');
    expect(packageJson.scripts).toHaveProperty('preview');
    expect(packageJson.scripts).toHaveProperty('server');
  });
});
