# Recently Finished Products API

## Overview

The Recently Finished Products API provides a dropdown for the `updatedetailsop` page that shows only products that are currently live (state 'ON') in the workpanel. This ensures that operators can only select products that are actively being processed.

## API Endpoints

### 1. GET `/api/live-products/recently-finished`

Returns a list of recently finished products that are currently live.

**Response Format:**
```json
{
  "recentlyFinishedProducts": [
    {
      "id": 1,
      "name": "Angel",
      "process": "Cutting MC/1",
      "state": "ON",
      "date": "12/7/2025",
      "createdAt": "2025-07-12T12:00:00.000Z",
      "machineName": "Cutting MC/1",
      "productName": "Angel",
      "displayName": "Angel - Cutting MC/1 (12/7/2025)",
      "availableCount": 3
    }
  ],
  "count": 1
}
```

**Key Features:**
- Only returns products with `state: 'ON'` (live products)
- Products are formatted as "ProductName - MachineName (Date) (3 available)"
- Sorted by creation date (most recent first)
- Includes availability count for each product

### 2. GET `/api/live-products/stream`

Provides Server-Sent Events (SSE) for real-time updates of live products.

**Event Types:**
- `connection_established`: Initial connection message
- `live_products_update`: Live products data update
- `error`: Error message if something goes wrong

## Frontend Integration

The `updatedetailsop` page has been updated to use the new API:

1. **Data Fetching**: Uses `/api/live-products/recently-finished` instead of the old endpoint
2. **Real-time Updates**: Connects to `/api/live-products/stream` for live updates
3. **Display Format**: Shows products in the format "ProductName - MachineName (Date) (3 available)"
4. **Filtering**: Only shows products that are currently live in the workpanel

## Database Schema

The API uses the existing database schema:

- **Job table**: Contains product processing information
- **Product table**: Contains product details
- **Machine table**: Contains machine information

## Key Benefits

1. **Consistency**: Only shows products that are actually live in the workpanel
2. **Real-time Updates**: Automatically updates when products are added/removed
3. **Clear Formatting**: Easy-to-read product names with machine and date information
4. **Availability Tracking**: Shows availability count for each product
5. **No Disturbance**: Doesn't affect any existing frontend or backend functionality

## Testing

Use the test script to verify functionality:

```bash
node scripts/test-recently-finished-products.js
```

This script:
- Creates a test product
- Verifies it appears in the API
- Moves it to past and verifies it's removed
- Tests the stream endpoint
- Cleans up test data

## Example Usage

```javascript
// Fetch recently finished products
const response = await fetch('/api/live-products/recently-finished');
const data = await response.json();

// Display in dropdown
data.recentlyFinishedProducts.forEach(product => {
  console.log(product.displayName); // "Angel - Cutting MC/1 (12/7/2025) (3 available)"
});
``` 