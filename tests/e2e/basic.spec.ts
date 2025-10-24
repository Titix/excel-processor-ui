import { test, expect } from '@playwright/test';

test.describe('Excel Processor - Basic E2E Test', () => {
  test('should load the application and display main elements', async ({ page }) => {
    // Navigate to the application
    await page.goto('/');
    
    // Wait for the page to load
    await page.waitForLoadState('networkidle');
    
    // Check if the main heading is visible
    await expect(page.locator('h1')).toContainText('Excel File Processor');
    
    // Check if the main button is visible
    await expect(page.locator('button:has-text("Select Folder")')).toBeVisible();
    
    // Check if the language selector is visible
    await expect(page.locator('text=Language:')).toBeVisible();
    
    // Check if the select element exists
    await expect(page.locator('select')).toBeVisible();
    
    // Check if footer is visible
    await expect(page.locator('text=© 2025 Excel Processor')).toBeVisible();
    
    // Verify the page loaded without critical errors
    const errors: string[] = [];
    page.on('console', msg => {
      if (msg.type() === 'error') {
        errors.push(msg.text());
      }
    });
    
    // Filter out non-critical errors
    const criticalErrors = errors.filter(error => 
      !error.includes('favicon') && 
      !error.includes('404') &&
      !error.includes('Failed to load resource')
    );
    
    expect(criticalErrors).toHaveLength(0);
  });
});
