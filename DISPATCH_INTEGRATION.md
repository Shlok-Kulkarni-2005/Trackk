# Dispatch Integration Documentation

## Overview

This document describes the new dispatch integration functionality that automatically adds products to the dispatched tab when they are updated with "Dispatched" status in the updatedetailsop page. The system now properly separates today's dispatches from dispatch history.

## Features

### 1. Automatic Dispatch Integration
- When a product is updated with "Dispatched" status in the updatedetailsop page, it automatically appears in the dispatched tab
- If the product is not present in dispatched items, it gets added with quantity 1
- If the product is already present, its quantity gets increased by 1

### 2. Real-time Quantity Tracking
- Products are grouped by name in the dispatched items
- Quantities are automatically calculated based on the number of times a product has been dispatched
- The dispatched tab shows accurate quantity counts for each product

### 3. Today vs History Separation
- **Today's Dispatches**: Shows only dispatches from the current date
- **Dispatch History**: Shows all dispatches (including today's)
- Automatic date-based filtering ensures proper separation
- Real-time updates when new dispatches are added

### 4. Enhanced Dispatched Tab
- Fetches data from the API instead of using static data
- Shows loading states while fetching data
- Includes a refresh button to manually update the data
- Displays proper quantities for each dispatched product
- Separate tabs for today's dispatches and dispatch history

## API Endpoints

### 1. `/api/dispatched/update-quantity`

**POST** - Add or update dispatched product quantity
```json
{
  "product": "Product Name",
  "quantity": 1,
  "operatorId": 12345
}
```

**Response:**
```json
{
  "success": true,
  "message": "Product added to dispatched items",
  "dispatchedItem": { ... }
}
```

**GET** - Get dispatched product quantity for a specific product
```
/api/dispatched/update-quantity?product=Product%20Name
```

**Response:**
```json
{
  "success": true,
  "product": "Product Name",
  "quantity": 2
}
```

### 2. `/api/admin/dispatched-items` (Enhanced)

**GET** - Fetch all dispatched items with grouped quantities and date separation
```json
{
  "success": true,
  "dispatchedItems": [
    {
      "id": "123",
      "product": "Product A",
      "quantity": 3,
      "destination": "Dispatched",
      "notes": "",
      "dispatchedAt": "2025-07-12T12:44:35.460Z",
      "status": "Dispatched",
      "totalQuantity": 3,
      "lastUpdated": "2025-07-12T12:44:35.460Z",
      "date": "2025-07-12"
    }
  ],
  "todayDispatchedItems": [
    {
      "id": "124",
      "product": "Product B",
      "quantity": 1,
      "destination": "Dispatched",
      "notes": "",
      "dispatchedAt": "2025-07-12T14:30:15.123Z",
      "status": "Dispatched",
      "totalQuantity": 1,
      "lastUpdated": "2025-07-12T14:30:15.123Z",
      "date": "2025-07-12"
    }
  ]
}
```

## Database Schema

The functionality uses the existing `OperatorProductUpdate` table:

```sql
model OperatorProductUpdate {
  id             Int      @id @default(autoincrement())
  operatorId     Int
  operator       Operator @relation(fields: [operatorId], references: [id])
  product        String
  processSteps   Json
  dispatchStatus String
  dispatchedCost Float
  createdAt      DateTime @default(now())
  archived       Boolean  @default(false)
}
```

## Integration Flow

### 1. Product Update Process
1. User selects a product in the updatedetailsop page
2. User sets dispatch status to "Dispatched"
3. User clicks update button
4. The `/api/operator/update` endpoint processes the update
5. If dispatch status is "Dispatched", it automatically calls `/api/dispatched/update-quantity`
6. The product is added to dispatched items or its quantity is increased

### 2. Dispatched Items Display
1. The dispatched tab fetches data from `/api/admin/dispatched-items`
2. The API groups products by name and calculates total quantities
3. The API separates today's dispatches from historical dispatches
4. The frontend displays the appropriate data based on the selected tab
5. Users can refresh the data manually using the refresh button

### 3. Date-Based Separation
1. The API calculates today's date in YYYY-MM-DD format
2. Items are filtered based on their `createdAt` date
3. Today's dispatches include only items created today
4. Dispatch history includes all dispatched items
5. Real-time updates maintain proper date separation

## Files Modified/Created

### New Files
- `app/api/dispatched/update-quantity/route.ts` - New API endpoint for managing dispatched quantities
- `scripts/test-dispatch-integration.js` - Database integration test
- `scripts/test-dispatch-simple.js` - API endpoint test
- `scripts/test-today-vs-history.js` - Today vs history separation test
- `scripts/test-multi-date-dispatch.js` - Multi-date dispatch test
- `DISPATCH_INTEGRATION.md` - This documentation file

### Modified Files
- `app/api/operator/update/route.ts` - Added automatic dispatch integration
- `app/api/admin/dispatched-items/route.ts` - Enhanced to group products, calculate quantities, and separate by date
- `app/dispatched/page.tsx` - Updated to fetch from API, show proper quantities, and handle today vs history separation

## Testing

### Database Integration Test
```bash
node scripts/test-dispatch-integration.js
```

### API Endpoint Test
```bash
node scripts/test-dispatch-simple.js
```

### Today vs History Separation Test
```bash
node scripts/test-today-vs-history.js
```

### Multi-Date Dispatch Test
```bash
node scripts/test-multi-date-dispatch.js
```

## Usage Examples

### 1. Updating a Product to Dispatched Status
1. Go to the updatedetailsop page
2. Select a product from the dropdown
3. Set dispatch status to "Dispatched"
4. Click "Update Details"
5. The product will automatically appear in the dispatched tab under "Today's Dispatched"

### 2. Viewing Today's Dispatches
1. Go to the dispatched tab
2. Click on "Today's Dispatched" tab
3. View only the dispatches from today
4. See accurate quantities for each product dispatched today

### 3. Viewing Dispatch History
1. Go to the dispatched tab
2. Click on "Dispatched History" tab
3. View all dispatched products from all dates
4. See total quantities for each product across all time

### 4. Checking Product Quantities
- Each product shows its total dispatched quantity
- Products are grouped by name, so multiple dispatches of the same product show as one entry with the total quantity
- Today's tab shows quantities for today only
- History tab shows quantities across all time

## Error Handling

- If the automatic dispatch integration fails, it doesn't affect the main product update
- API errors are logged but don't crash the application
- The frontend shows loading states and error messages when appropriate
- Date filtering gracefully handles timezone differences

## Future Enhancements

1. **Cost Tracking**: Add cost calculation for dispatched items
2. **Date Range Filtering**: Add custom date range filters for dispatch history
3. **Export Functionality**: Add ability to export dispatched items to CSV/Excel
4. **Real-time Updates**: Add WebSocket support for real-time updates
5. **Bulk Operations**: Add ability to dispatch multiple products at once
6. **Advanced Filtering**: Add filters by operator, machine, or product type

## Troubleshooting

### Common Issues

1. **Product not appearing in today's dispatches**
   - Check if the dispatch status was set to "Dispatched"
   - Verify the API call was successful in browser developer tools
   - Check if the dispatch was created today (not yesterday or earlier)
   - Try refreshing the dispatched tab

2. **Quantity not updating**
   - Check if the product name matches exactly (case-sensitive)
   - Verify the database has the correct records
   - Run the test scripts to verify functionality

3. **Date separation not working**
   - Check if the server timezone is correct
   - Verify the date format in the database
   - Run the multi-date test script to verify separation

4. **API errors**
   - Check if the development server is running
   - Verify database connection
   - Check browser console for error messages

### Debug Commands

```bash
# Test database integration
node scripts/test-dispatch-integration.js

# Test API endpoints
node scripts/test-dispatch-simple.js

# Test today vs history separation
node scripts/test-today-vs-history.js

# Test multi-date dispatch
node scripts/test-multi-date-dispatch.js

# Check database directly
npx prisma studio
```

## Conclusion

The dispatch integration provides a seamless workflow for tracking dispatched products with proper date-based separation. When operators update products with "Dispatched" status, the system automatically manages the dispatched items list, ensuring accurate quantity tracking, real-time updates, and proper separation between today's dispatches and historical data. 