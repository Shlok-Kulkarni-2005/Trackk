'use client';
import React, { useState, useEffect } from 'react';

export default function TestProcountPage() {
  const [data, setData] = useState<any>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const fetchData = async () => {
      try {
        setLoading(true);
        console.log('Testing procount API...');
        const response = await fetch('/api/procount');
        console.log('Response:', response);
        
        if (response.ok) {
          const result = await response.json();
          console.log('API Result:', result);
          setData(result);
        } else {
          setError(`HTTP ${response.status}: ${response.statusText}`);
        }
      } catch (err) {
        console.error('Error:', err);
        setError(err instanceof Error ? err.message : 'Unknown error');
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  return (
    <div className="p-8">
      <h1 className="text-2xl font-bold mb-4">Procount API Test</h1>
      
      {loading && <p>Loading...</p>}
      
      {error && (
        <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded mb-4">
          <strong>Error:</strong> {error}
        </div>
      )}
      
      {data && (
        <div>
          <h2 className="text-xl font-semibold mb-2">API Response:</h2>
          <pre className="bg-gray-100 p-4 rounded overflow-auto">
            {JSON.stringify(data, null, 2)}
          </pre>
          
          <h2 className="text-xl font-semibold mt-4 mb-2">Product Counts:</h2>
          {data.productCounts && data.productCounts.length > 0 ? (
            <ul className="space-y-2">
              {data.productCounts.map((item: any) => (
                <li key={item.id} className="bg-white p-3 rounded border">
                  <strong>{item.product.name}</strong> - Count: {item.count}
                  <br />
                  <small className="text-gray-600">
                    Machine: {item.machine} | State: {item.state}
                  </small>
                </li>
              ))}
            </ul>
          ) : (
            <p className="text-gray-500">No products found</p>
          )}
        </div>
      )}
    </div>
  );
} 