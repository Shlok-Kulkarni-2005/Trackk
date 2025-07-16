'use client';
import React, { useState, useMemo, useEffect } from 'react';
import { Menu, X, Package, Calendar, Search, RefreshCw } from 'lucide-react';
import Sidebar from '../../components/sidebarm';

// Interface for dispatched item data
interface DispatchedItem {
  id: number;
  productId: string;
  product: string;
  quantity: number;
  cost: number;
  date: string;
  destination?: string;
  notes?: string;
  type?: string;
}

export default function DispatchedPage(): React.ReactElement {
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const [activeSection, setActiveSection] = useState<'today' | 'history'>('today');
  const [searchTermRFD, setSearchTermRFD] = useState<string>('');

  // Dispatched items data from API
  const [dispatchedItems, setDispatchedItems] = useState<DispatchedItem[]>([]);
  const [todayDispatchedItems, setTodayDispatchedItems] = useState<DispatchedItem[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);

  // Fetch dispatched items from API
  useEffect(() => {
    async function fetchDispatchedItems() {
      setIsLoading(true);
      try {
        const res = await fetch('/api/admin/dispatched-items');
        if (res.ok) {
          const data = await res.json();
          if (data.success) {
            // Transform API data to match the interface
            const transformedItems = data.dispatchedItems.map((item: any, index: number) => ({
              id: item.id.startsWith('rfd_') ? index + 10000 : parseInt(item.id) || index + 20000,
              productId: item.product, // Using product name as ID for now
              product: item.product,
              quantity: item.quantity || 1,
              cost: 0, // Default cost since it's not provided by API
              date: new Date(item.dispatchedAt).toLocaleDateString('en-CA'),
              destination: item.destination,
              notes: item.notes,
              type: item.type
            }));
            
            const todayTransformedItems = data.todayDispatchedItems.map((item: any, index: number) => ({
              id: item.id.startsWith('rfd_') ? index + 30000 : parseInt(item.id) || index + 40000,
              productId: item.product,
              product: item.product,
              quantity: item.quantity || 1,
              cost: 0,
              date: new Date(item.dispatchedAt).toLocaleDateString('en-CA'),
              destination: item.destination,
              notes: item.notes,
              type: item.type
            }));
            
            setDispatchedItems(transformedItems);
            setTodayDispatchedItems(todayTransformedItems);
          } else {
            console.error('Failed to fetch dispatched items:', data.error);
          }
        } else {
          console.error('Failed to fetch dispatched items:', res.status);
        }
      } catch (error) {
        console.error('Error fetching dispatched items:', error);
      } finally {
        setIsLoading(false);
      }
    }
    fetchDispatchedItems();

    // Set up real-time polling every 30 seconds
    const interval = setInterval(fetchDispatchedItems, 30000);

    // Cleanup interval on component unmount
    return () => clearInterval(interval);
  }, []);

  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };

  const handleRefresh = async (): Promise<void> => {
    setIsLoading(true);
    try {
      const res = await fetch('/api/admin/dispatched-items');
      if (res.ok) {
        const data = await res.json();
        if (data.success) {
          const transformedItems = data.dispatchedItems.map((item: any, index: number) => ({
            id: item.id.startsWith('rfd_') ? index + 10000 : parseInt(item.id) || index + 20000,
            productId: item.product,
            product: item.product,
            quantity: item.quantity || 1,
            cost: 0,
            date: new Date(item.dispatchedAt).toLocaleDateString('en-CA'),
            destination: item.destination,
            notes: item.notes,
            type: item.type
          }));
          
          const todayTransformedItems = data.todayDispatchedItems.map((item: any, index: number) => ({
            id: item.id.startsWith('rfd_') ? index + 30000 : parseInt(item.id) || index + 40000,
            productId: item.product,
            product: item.product,
            quantity: item.quantity || 1,
            cost: 0,
            date: new Date(item.dispatchedAt).toLocaleDateString('en-CA'),
            destination: item.destination,
            notes: item.notes,
            type: item.type
          }));
          
          setDispatchedItems(transformedItems);
          setTodayDispatchedItems(todayTransformedItems);
        }
      }
    } catch (error) {
      console.error('Error refreshing dispatched items:', error);
    } finally {
      setIsLoading(false);
    }
  };

  // Filtered lists for RFD only
  const todayRFD = todayDispatchedItems.filter(item => item.type === 'rfd');
  const historyRFD = dispatchedItems.filter(item => item.type === 'rfd');

  // Search logic
  const filteredTodayRFD = todayRFD.filter(item => item.product.toLowerCase().includes(searchTermRFD.toLowerCase()));
  const filteredHistoryRFD = historyRFD.filter(item => item.product.toLowerCase().includes(searchTermRFD.toLowerCase()));

  // Calculate summary values for today's RFD
  const totalTodayRFDQuantity = filteredTodayRFD.reduce((sum, item) => sum + (item.quantity || 0), 0);
  const uniqueTodayRFDProducts = new Set(filteredTodayRFD.map(item => item.product)).size;

  // Calculate summary values for RFD history
  const totalHistoryRFDQuantity = filteredHistoryRFD.reduce((sum, item) => sum + (item.quantity || 0), 0);
  const uniqueHistoryRFDProducts = new Set(filteredHistoryRFD.map(item => item.product)).size;

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Sidebar Component */}
      <Sidebar 
        isOpen={sidebarOpen} 
        onClose={() => setSidebarOpen(false)} 
        username={null}
      />

      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-4 py-3">
        <div className="flex items-center justify-between">
          <button 
            onClick={() => setSidebarOpen(true)}
            className="p-2 hover:bg-gray-100 rounded-lg"
            type="button"
            aria-label="Open menu"
          >
            <Menu className="w-6 h-6 text-gray-600" />
          </button>
          <h1 className="text-lg font-semibold text-gray-900">RFD Status</h1>
          <div className="flex items-center space-x-2">
            <button
              onClick={handleRefresh}
              disabled={isLoading}
              className="p-2 hover:bg-gray-100 rounded-lg disabled:opacity-50"
              type="button"
              aria-label="Refresh data"
            >
              <RefreshCw className={`w-5 h-5 text-gray-600 ${isLoading ? 'animate-spin' : ''}`} />
            </button>
            <div className="w-8 h-8 bg-indigo-600 rounded-full flex items-center justify-center">
              <span className="text-white font-medium text-sm">A</span>
            </div>
          </div>
        </div>
      </header>

      {/* Section Tabs: Today / History */}
      <div className="flex border-b">
        <button
          onClick={() => setActiveSection('today')}
          className={`flex-1 py-3 flex items-center justify-center ${
            activeSection === 'today' 
              ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-700' 
              : 'text-gray-600 hover:bg-gray-100'
          }`}
        >
          Today&apos;s RFD
        </button>
        <button
          onClick={() => setActiveSection('history')}
          className={`flex-1 py-3 flex items-center justify-center ${
            activeSection === 'history' 
              ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-700' 
              : 'text-gray-600 hover:bg-gray-100'
          }`}
        >
          RFD History
        </button>
      </div>

      {/* Main Content */}
      <main className="flex-1 px-4 py-6">
        {/* Search Bar */}
        <div className="mb-6">
          <div className="relative">
            <Search className="absolute left-3 top-3 w-5 h-5 text-gray-400" />
            <input
              type="text"
              placeholder={'Search RFD...'}
              value={searchTermRFD}
              onChange={(e) => setSearchTermRFD(e.target.value)}
              className="w-full pl-10 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 text-sm"
              aria-label={'Search RFD'}
            />
          </div>
        </div>

        {/* Summary Cards (only for today/history) */}
        {activeSection === 'today' && (
          <div className="grid grid-cols-2 gap-4 mb-6">
            <div className="bg-white rounded-lg p-4 shadow-sm border border-gray-200">
              <div className="flex items-center space-x-3">
                <div className="w-8 h-8 bg-green-100 rounded-lg flex items-center justify-center">
                  <Package className="w-4 h-4 text-green-600" />
                </div>
                <div>
                  <p className="text-xs text-gray-500">Today&apos;s RFD (Total Quantity)</p>
                  <p className="text-lg font-bold text-gray-900">{totalTodayRFDQuantity}</p>
                </div>
              </div>
            </div>
            <div className="bg-white rounded-lg p-4 shadow-sm border border-gray-200">
              <div className="flex items-center space-x-3">
                <div className="w-8 h-8 bg-blue-100 rounded-lg flex items-center justify-center">
                  <Package className="w-4 h-4 text-blue-600" />
                </div>
                <div>
                  <p className="text-xs text-gray-500">Today&apos;s Unique Items</p>
                  <p className="text-lg font-bold text-gray-900">{uniqueTodayRFDProducts}</p>
                </div>
              </div>
            </div>
          </div>
        )}
        {activeSection === 'history' && (
          <div className="grid grid-cols-2 gap-4 mb-6">
            <div className="bg-white rounded-lg p-4 shadow-sm border border-gray-200">
              <div className="flex items-center space-x-3">
                <div className="w-8 h-8 bg-green-100 rounded-lg flex items-center justify-center">
                  <Package className="w-4 h-4 text-green-600" />
                </div>
                <div>
                  <p className="text-xs text-gray-500">RFD History (Total Quantity)</p>
                  <p className="text-lg font-bold text-gray-900">{totalHistoryRFDQuantity}</p>
                </div>
              </div>
            </div>
            <div className="bg-white rounded-lg p-4 shadow-sm border border-gray-200">
              <div className="flex items-center space-x-3">
                <div className="w-8 h-8 bg-blue-100 rounded-lg flex items-center justify-center">
                  <Package className="w-4 h-4 text-blue-600" />
                </div>
                <div>
                  <p className="text-xs text-gray-500">History Unique Items</p>
                  <p className="text-lg font-bold text-gray-900">{uniqueHistoryRFDProducts}</p>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Items List */}
        <div>
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-base font-semibold text-gray-900">
              {activeSection === 'today' ? "Today&apos;s RFD" : 'RFD History'}
            </h3>
            <span className="text-sm text-gray-500">
              {activeSection === 'today' ? filteredTodayRFD.length : filteredHistoryRFD.length} items
            </span>
          </div>

          {isLoading ? (
            <div className="text-center py-8">
              <div className="text-gray-500">Loading items...</div>
            </div>
          ) : (
            <div className="space-y-4">
              {activeSection === 'today' && filteredTodayRFD.map((item) => (
                <div key={item.id} className="bg-white rounded-lg p-4 shadow-sm border border-gray-200">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-3">
                      <div className="w-10 h-10 bg-indigo-100 rounded-lg flex items-center justify-center">
                        <Package className="w-5 h-5 text-indigo-600" />
                      </div>
                      <div>
                        <h4 className="text-sm font-medium text-gray-900">{item.product}</h4>
                        <div className="flex items-center space-x-4 text-xs text-gray-500 mt-1">
                          <span className="flex items-center space-x-1">
                            <Calendar className="w-3 h-3" />
                            <span>{item.date}</span>
                          </span>
                        </div>
                        <div className="text-sm font-medium text-gray-900">Qty: {item.quantity}</div>
                      </div>
                    </div>
                  </div>
                </div>
              ))}
              {activeSection === 'history' && filteredHistoryRFD.map((item) => (
                <div key={item.id} className="bg-white rounded-lg p-4 shadow-sm border border-gray-200">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-3">
                      <div className="w-10 h-10 bg-indigo-100 rounded-lg flex items-center justify-center">
                        <Package className="w-5 h-5 text-indigo-600" />
                      </div>
                      <div>
                        <h4 className="text-sm font-medium text-gray-900">{item.product}</h4>
                        <div className="flex items-center space-x-4 text-xs text-gray-500 mt-1">
                          <span className="flex items-center space-x-1">
                            <Calendar className="w-3 h-3" />
                            <span>{item.date}</span>
                          </span>
                        </div>
                        <div className="text-sm font-medium text-gray-900">Qty: {item.quantity}</div>
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </main>
    </div>
  );
}