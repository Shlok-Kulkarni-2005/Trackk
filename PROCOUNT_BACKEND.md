# Procount Backend Documentation

## Overview
The Procount backend tracks products when the RFD stage is selected and status is OFF in the Add Jobs tab. It maintains a count of how many times each product has been selected under these specific conditions.

## Database Schema

### ProductCount Model
```prisma
model ProductCount {
  id        Int      @id @default(autoincrement())
  product   Product  @relation(fields: [productId], references: [id])
  productId Int
  count     Int      @default(0)
  machine   String
  stage     String
  state     String
  createdAt DateTime @default(now())
  updatedAt DateTime @updatedAt

  @@unique([productId])
}
```

## API Endpoints

### GET /api/procount
Fetches all product counts with their associated product information.

**Response:**
```json
{
  "productCounts": [
    {
      "id": 1,
      "product": {
        "id": 1,
        "name": "Product A"
      },
      "count": 5,
      "machine": "RFD",
      "stage": "RFD",
      "state": "OFF",
      "createdAt": "2024-01-01T00:00:00Z",
      "updatedAt": "2024-01-01T00:00:00Z"
    }
  ]
}
```

### POST /api/procount
Updates product count when RFD stage is selected and status is OFF.

**Request Body:**
```json
{
  "productName": "Product A",
  "machine": "RFD",
  "stage": "RFD",
  "state": "OFF"
}
```

**Response:**
```json
{
  "productCount": {
    "id": 1,
    "product": {
      "id": 1,
      "name": "Product A"
    },
    "count": 6,
    "machine": "RFD",
    "stage": "RFD",
    "state": "OFF",
    "createdAt": "2024-01-01T00:00:00Z",
    "updatedAt": "2024-01-01T00:00:00Z"
  },
  "message": "Product count updated successfully"
}
```

### DELETE /api/procount
Clears all product count records.

## Real-time Updates

### GET /api/procount/stream
Server-Sent Events endpoint for real-time updates.

**Events:**
- `productCountUpdated`: Sent when a product count is updated

**Event Data:**
```json
{
  "type": "productCountUpdated",
  "productCount": {
    "id": 1,
    "name": "Product A",
    "count": 6,
    "status": "OFF",
    "machine": "RFD",
    "stage": "RFD",
    "state": "OFF",
    "updatedAt": "2024-01-01T00:00:00Z"
  }
}
```

## Integration with Jobs API

The jobs API (`/api/jobs`) automatically calls the procount API when:
- Stage is "RFD"
- State is "OFF"

This ensures that product counts are automatically updated whenever these conditions are met in the Add Jobs form.

## Frontend Integration

The procount page (`/procount`) automatically:
1. Fetches initial product counts on load
2. Establishes real-time connection for live updates
3. Updates the UI when new product counts are received
4. Shows loading and error states appropriately

## Testing

Use the test script `scripts/test-procount-backend.js` to verify the backend functionality:

```bash
node scripts/test-procount-backend.js
```

This script tests:
1. Job creation with RFD stage and OFF status
2. Count increment for the same product
3. Jobs with different stages (should not affect count)
4. Fetching all product counts
5. Direct procount API calls

## Key Features

1. **Automatic Tracking**: Product counts are automatically updated when RFD stage and OFF status are selected
2. **Real-time Updates**: Frontend receives live updates via Server-Sent Events
3. **Unique Products**: Each product can only have one count record (enforced by database constraint)
4. **Incremental Counting**: Existing products have their count incremented, new products start at 1
5. **Error Handling**: Comprehensive error handling and validation
6. **Loading States**: Frontend shows appropriate loading and error states 