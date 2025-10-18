/**
 * Global Test Teardown
 * This file runs once after all tests
 */

// Clean up global resources
afterAll(() => {
  // Clear any remaining timers
  jest.clearAllTimers();
  
  // Clear any remaining mocks
  jest.clearAllMocks();
  
  // Reset DOM
  if (typeof document !== 'undefined') {
    document.body.innerHTML = '';
  }
  
  // Clear storage
  if (typeof localStorage !== 'undefined') {
    localStorage.clear();
  }
  
  if (typeof sessionStorage !== 'undefined') {
    sessionStorage.clear();
  }
  
  // Clear any global variables
  if (typeof global !== 'undefined') {
    delete (global as any).testUtils;
  }
});

console.log('🧹 Test environment cleaned up');
