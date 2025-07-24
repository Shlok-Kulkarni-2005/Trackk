'use client';
import React, { useState, useEffect, useRef } from 'react';
import { Menu, ChevronDown, X } from 'lucide-react';
import Sidebar from '../../components/sidebarm';


// Type definitions
interface Machine {
  id: number;
  name: string;
  status: 'ON' | 'OFF';
  statusColor: string;
}

interface Product {
  id: string;
  name: string;
  operation: string;
  date: string;
  expiryDate?: string; // Optional field to track product lifecycle
  state: 'ON' | 'OFF';
  quantity: number; // Added for quantity display
}


type FilterType = 'Machine/Process No' | 'Product Type';

interface CustomDropdownProps {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
}


export default function WorkPanelInterface() {
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const [currentView, setCurrentView] = useState<'live' | 'past'>('live');
  const [selectedProduct, setSelectedProduct] = useState<Product | null>(null);

  const [productData, setProductData] = useState<Product[]>([]);
  const [allJobs, setAllJobs] = useState<any[]>([]);
  const [now, setNow] = useState(Date.now());
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' | 'info' } | null>(null);

  // Auto-hide messages after 5 seconds
  useEffect(() => {
    if (message) {
      const timer = setTimeout(() => {
        setMessage(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [message]);

  // Update timer every second for live products
  useEffect(() => {
    const interval = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    async function fetchJobs() {
      const res = await fetch('/api/jobs');
      if (res.ok) {
        const data = await res.json();
        setAllJobs(data.jobs);
        // Group jobs by productId, machineName, and state for accurate quantity
        const jobsByKey: { [key: string]: any[] } = {};
        data.jobs.forEach((job: any) => {
          const key = `${job.product.id}__${job.machine.name}__${job.state}`;
          if (!jobsByKey[key]) jobsByKey[key] = [];
          jobsByKey[key].push(job);
        });
        // Build productData with ON and OFF quantities for each (product, machine) pair
        const products: Product[] = [];
        Object.entries(jobsByKey).forEach(([key, jobs]) => {
          const [productId, machineName, state] = key.split('__');
          const productName = jobs[0].product.name;
          const date = new Date(jobs[0].createdAt).toLocaleDateString();
          products.push({
            id: `${productId}__${machineName}__${state}`,
            name: productName,
            operation: machineName,
            date,
            state: state as 'ON' | 'OFF',
            quantity: jobs.length,
          });
        });
        setProductData(products);
      }
    }
    fetchJobs();
  }, []);

  // Helper to get all jobs for a product-machine pair
  const getJobsForProduct = (productId: string, operation: string, state: 'ON' | 'OFF') => {
    // productId is now like "123__Cutting MC/1__ON"
    const [pid, machineName] = productId.split('__');
    return allJobs.filter(job => job.product.id.toString() === pid && job.machine.name === operation && job.state === state)
      .sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
  };

  // Helper to format duration as HH:mm:ss
  const formatDuration = (ms: number) => {
    const totalSeconds = Math.floor(ms / 1000);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
  };



  // Machine/Process data
  const machineData: Machine[] = [
    { id: 1, name: 'Cutting MC/1', status: 'ON', statusColor: 'green' },
    { id: 2, name: 'Milling 1', status: 'OFF', statusColor: 'gray' },
    { id: 3, name: 'Milling 2', status: 'ON', statusColor: 'green' },
    { id: 4, name: 'Drilling', status: 'OFF', statusColor: 'gray' },
    { id: 5, name: 'CNC Finish', status: 'ON', statusColor: 'green' }
  ];



  // Ensure live and past products are mutually exclusive
  const getLiveProducts = (): Product[] => {
    return productData.filter(product => product.state === 'ON' && product.quantity > 0);
  };

  const getPastProducts = (): Product[] => {
    return productData.filter(product => product.state === 'OFF' && product.quantity > 0);
  };



  const getStatusColor = (status: Machine['status']): string => {
    switch (status) {
      case 'ON': return 'text-green-600';
      case 'OFF': return 'text-gray-500';
      default: return 'text-gray-500';
    }
  };

  const getStatusDotColor = (status: Machine['status']): string => {
    switch (status) {
      case 'ON': return 'bg-green-500';
      case 'OFF': return 'bg-gray-400';
      default: return 'bg-gray-400';
    }
  };

  const handleMachineClick = (machine: Machine): void => {
    console.log('Machine clicked:', machine);
  };

  const handleSeeDetails = (product: Product): void => {
    setSelectedProduct(product);
  };

  const handleClose = (): void => {
    setSelectedProduct(null);
  };

  const handleMoveToPast = async (product: Product): Promise<void> => {
    try {
      // Find the job ID for this product
      const res = await fetch('/api/jobs');
      if (res.ok) {
        const data = await res.json();
        const productJob = data.jobs.find((job: any) => 
          job.product.id.toString() === product.id && job.state === 'ON'
        );
        
        if (productJob) {
          const moveRes = await fetch('/api/products/lifecycle', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              productId: product.id,
              jobId: productJob.id,
              action: 'move_to_past',
              reason: 'manually_moved_from_workpanel'
            })
          });

          if (moveRes.ok) {
            setMessage({ text: 'Product moved to past products successfully!', type: 'success' });
            // Refresh the product data
            const refreshRes = await fetch('/api/jobs');
            if (refreshRes.ok) {
              const refreshData = await refreshRes.json();
              const jobsByKey: { [key: string]: any[] } = {};
              refreshData.jobs.forEach((job: any) => {
                const key = `${job.product.id}__${job.machine.name}__${job.state}`;
                if (!jobsByKey[key]) jobsByKey[key] = [];
                jobsByKey[key].push(job);
              });
              const products: Product[] = [];
              Object.entries(jobsByKey).forEach(([key, jobs]) => {
                const [productId, machineName, state] = key.split('__');
                const productName = jobs[0].product.name;
                const date = new Date(jobs[0].createdAt).toLocaleDateString();
                products.push({
                  id: `${productId}__${machineName}__${state}`,
                  name: productName,
                  operation: machineName,
                  date,
                  state: state as 'ON' | 'OFF',
                  quantity: jobs.length,
                });
              });
              setProductData(products);
            }
            setSelectedProduct(null);
          } else {
            setMessage({ text: 'Failed to move product to past', type: 'error' });
          }
        } else {
          setMessage({ text: 'Product job not found', type: 'info' });
        }
      }
    } catch (error) {
      console.error('Error moving product to past:', error);
      setMessage({ text: 'Failed to move product to past', type: 'error' });
    }
  };


    const CustomDropdown = ({ label, value, options, onChange }: CustomDropdownProps) => {
    const [isOpen, setIsOpen] = useState<boolean>(false);
    const dropdownRef = useRef<HTMLDivElement>(null);

    // Handle click outside dropdown to close it
    useEffect(() => {
      function handleClickOutside(event: MouseEvent) {
        if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
          setIsOpen(false);
        }
      }

      if (isOpen) {
        document.addEventListener('mousedown', handleClickOutside);
      }

      return () => {
        document.removeEventListener('mousedown', handleClickOutside);
      };
    }, [isOpen]);

    const handleOptionSelect = (option: string, event: React.MouseEvent) => {
      event.preventDefault();
      event.stopPropagation();
      console.log('Workpanel dropdown: Option selected:', option);
      onChange(option);
      setIsOpen(false);
    };

    return (
      <div className="mb-6">
        <label className="block text-gray-700 font-medium mb-3 text-sm">
          {label}
        </label>
        <div className="relative" ref={dropdownRef}>
          <button
            type="button"
            onClick={() => {
              console.log('Workpanel dropdown clicked, current state:', isOpen);
              console.log('Available options:', options);
              setIsOpen(!isOpen);
            }}
            className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-left flex items-center justify-between hover:border-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
          >
            <span className="text-gray-700 text-sm">{value}</span>
            <div className="flex items-center space-x-2">
              <ChevronDown 
                className={`w-4 h-4 text-gray-400 transition-transform duration-200 ${
                  isOpen ? 'rotate-180' : ''
                }`} 
              />
            </div>
          </button>
          
          {isOpen && (
            <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-gray-300 rounded-lg shadow-xl z-50 max-h-48 overflow-y-auto">
              {options.length === 0 ? (
                <div className="px-4 py-3 text-gray-500 text-center">
                  <div className="text-sm">No options available</div>
                </div>
              ) : (
                options.map((option, index) => (
                  <button
                    key={index}
                    type="button"
                    onClick={(e) => handleOptionSelect(option, e)}
                    className={`w-full px-4 py-3 text-left hover:bg-blue-50 text-gray-700 text-sm first:rounded-t-lg last:rounded-b-lg transition-colors focus:outline-none focus:bg-blue-50 ${
                      value === option ? 'bg-blue-50 border-blue-200' : ''
                    }`}
                  >
                    {option}
                    {value === option && (
                      <span className="text-xs text-blue-600 ml-2">✓ Selected</span>
                    )}
                  </button>
                ))
              )}
            </div>
          )}
        </div>
      </div>
    );
  };

  const renderMachineView = () => (
    <div className="space-y-4">
      {machineData.map((machine) => (
        <button
          key={machine.id}
          onClick={() => handleMachineClick(machine)}
          className="w-full bg-white rounded-lg p-4 shadow-sm border border-gray-200 hover:shadow-md hover:border-gray-300 transition-all text-left"
        >
          <div className="flex items-center space-x-4">
            <div className="flex-shrink-0">
              <div className={`w-3 h-3 rounded-full ${getStatusDotColor(machine.status)}`}></div>
            </div>
            <div className="flex-1">
              <h3 className="text-gray-900 font-medium text-base mb-1">
                {machine.name}
              </h3>
              <p className={`text-sm font-medium ${getStatusColor(machine.status)}`}>
                {machine.status}
              </p>
            </div>
          </div>
        </button>
      ))}
      
      {/* Status Overview */}
      <div className="mt-8 bg-white rounded-lg p-6 shadow-sm border border-gray-200">
        <h3 className="text-gray-800 font-medium mb-6">Status Overview</h3>
        <div className="grid grid-cols-2 gap-8">
          <div className="text-center">
            <div className="text-4xl font-bold text-green-600 mb-2">
              {machineData.filter(m => m.status === 'ON').length}
            </div>
            <div className="text-sm text-gray-600 flex items-center justify-center">
              <div className="w-2 h-2 bg-green-500 rounded-full mr-2"></div>
              Online
            </div>
          </div>
          <div className="text-center">
            <div className="text-4xl font-bold text-gray-500 mb-2">
              {machineData.filter(m => m.status === 'OFF').length}
            </div>
            <div className="text-sm text-gray-600 flex items-center justify-center">
              <div className="w-2 h-2 bg-gray-400 rounded-full mr-2"></div>
              Offline
            </div>
          </div>
        </div>
      </div>
    </div>
  );

  // Update renderProductList to use new getJobsForProduct signature
  const renderProductList = (products: Product[]) => (
    <div className="space-y-4">
      {products.map((product, idx) => {
        const jobs = getJobsForProduct(product.id, product.operation, product.state);
        let timingInfo = '';
        if (product.state === 'ON') {
          const lastOnJob = [...jobs].reverse().find(j => j.state === 'ON');
          if (lastOnJob) {
            const start = new Date(lastOnJob.createdAt).getTime();
            timingInfo = `Running: ${formatDuration(now - start)}`;
          }
        } else {
          // For past products: find ON job and OFF job, calculate duration
          const sortedJobs = [...jobs].sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
          if (jobs.length === 1 && jobs[0].state === 'OFF') {
            const job = jobs[0];
            const duration = new Date(job.updatedAt || job.createdAt).getTime() - new Date(job.createdAt).getTime();
            if (duration > 0) {
              timingInfo = `Final: ${formatDuration(duration)}`;
            } else {
              timingInfo = `Completed`;
            }
          }
          const lastOffJob = sortedJobs.reverse().find(j => j.state === 'OFF');
          if (lastOffJob) {
            const jobsBeforeOff = sortedJobs.slice(sortedJobs.indexOf(lastOffJob) + 1);
            const lastOnJob = jobsBeforeOff.reverse().find(j => j.state === 'ON');
            if (lastOnJob) {
              const offTime = new Date(lastOffJob.updatedAt || lastOffJob.createdAt).getTime();
              const onTime = new Date(lastOnJob.createdAt).getTime();
              const duration = offTime - onTime;
              if (duration > 0) {
                timingInfo = `Final: ${formatDuration(duration)}`;
              } else {
                timingInfo = `Total: ${formatDuration(0)}`;
              }
            } else {
              timingInfo = `Completed`;
            }
          } else {
            timingInfo = `Completed`;
          }
        }
        // Use product.quantity directly
        return (
          <div key={product.id + '-' + idx} className="bg-white rounded-lg p-4 shadow-sm border border-gray-200">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-4">
                <div className={`w-3 h-3 rounded-full ${product.state === 'ON' ? 'bg-green-500' : 'bg-gray-400'}`}></div>
                <div>
                  <h3 className="text-gray-900 font-medium text-base mb-1">{product.name}</h3>
                  {timingInfo && <div className="text-xs text-gray-400 mb-1">{timingInfo}</div>}
                  <p className="text-sm text-gray-500">Operation: {product.operation}</p>
                  <p className="text-xs text-gray-400">Date: {product.date}</p>
                  <p className={`text-xs font-semibold ${product.state === 'ON' ? 'text-green-600' : 'text-gray-500'}`}>State: {product.state}</p>
                  <p className="text-xs font-semibold text-blue-600">Quantity: {product.quantity}</p>
                </div>
              </div>
              <button
                onClick={() => handleSeeDetails(product)}
                className="px-4 py-2 bg-black text-white text-sm rounded-md hover:bg-gray-800 transition-colors"
              >
                See Details
              </button>
            </div>
          </div>
        );
      })}
    </div>
  );

  const renderDetailsView = () => {
    if (!selectedProduct) return null;
    const jobs = getJobsForProduct(selectedProduct.id, selectedProduct.operation, selectedProduct.state);
    // Debug log to inspect jobs for the selected product
    console.log('Jobs for selected product:', jobs);
    // Calculate time using standard logic
    let timeInfo = '';
    

    
    if (selectedProduct.state === 'ON') {
      // For running products: now - last ON job createdAt
      const lastOnJob = [...jobs].reverse().find(j => j.state === 'ON');
      if (lastOnJob) {
        const start = new Date(lastOnJob.createdAt).getTime();
        timeInfo = `Running: ${formatDuration(now - start)}`;
      }
    } else {
      // For past products: use earliest createdAt and latest updatedAt
      if (jobs.length > 0) {
        const start = new Date(Math.min(...jobs.map(j => new Date(j.createdAt).getTime())));
        const end = new Date(Math.max(...jobs.map(j => new Date(j.updatedAt || j.createdAt).getTime())));
        const duration = end.getTime() - start.getTime();
        timeInfo = `Final: ${formatDuration(duration)}`;
      } else {
        timeInfo = `Completed`;
      }
    }
    // Calculate quantity
    let quantity = 0;
    if (selectedProduct.state === 'ON') {
      // Live product: count ON jobs
      quantity = jobs.filter(j => j.state === 'ON').length;
    } else {
      // Past product: count OFF jobs (since all ON jobs were turned OFF at once)
      quantity = jobs.filter(j => j.state === 'OFF').length;
    }
    return (
      <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center space-x-4">
            <div className="w-3 h-3 rounded-full bg-blue-500"></div>
            <h1 className="text-2xl font-bold text-gray-900">{selectedProduct?.name}</h1>
          </div>
          <div className="flex items-center space-x-2">
            <button
              onClick={handleClose}
              className="flex items-center space-x-2 px-4 py-2 bg-gray-600 text-white rounded-md hover:bg-gray-700 transition-colors"
            >
              <span>Close</span>
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
        <div className="space-y-6">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Date : DD/MM/YY</label>
            <div className="text-gray-900 text-lg">{selectedProduct?.date}</div>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Operation</label>
            <div className="text-gray-900 text-lg">{selectedProduct?.operation}</div>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Status</label>
            <div className="inline-flex items-center px-3 py-1 rounded-full text-sm bg-green-100 text-green-800">
              {selectedProduct.state === 'ON' ? 'Active' : 'Inactive'}
            </div>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Time</label>
            <div className="text-gray-900 text-lg">{timeInfo}</div>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Quantity</label>
            <div className="text-gray-900 text-lg">{quantity}</div>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gray-100">
      {/* Sidebar Component */}
      <Sidebar 
        isOpen={sidebarOpen} 
        onClose={() => setSidebarOpen(false)} 
        username={null}
      />

      {/* Header */}
      <header className="bg-white px-6 py-4 flex items-center justify-between sticky top-0 z-20">
        <button 
          onClick={() => setSidebarOpen(true)}
          className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
        >
          <Menu className="w-6 h-6 text-blue-700" />
        </button>
        
        <h1 className="text-xl font-semibold text-blue-700">Work Panel</h1>
        
        <div className="w-10 h-10 bg-blue-600 rounded-full flex items-center justify-center">
          <span className="text-white font-medium text-lg">A</span>
        </div>
      </header>

      {/* Section Tabs */}
      <div className="flex border-b">
        <button
          onClick={() => setCurrentView('live')}
          className={`flex-1 py-3 flex items-center justify-center ${
            currentView === 'live' 
              ? 'bg-blue-100 text-blue-700 border-b-2 border-blue-700' 
              : 'text-gray-600 hover:bg-gray-100'
          }`}
        >
          Live Products
        </button>
        <button
          onClick={() => setCurrentView('past')}
          className={`flex-1 py-3 flex items-center justify-center ${
            currentView === 'past' 
              ? 'bg-blue-100 text-blue-700 border-b-2 border-blue-700' 
              : 'text-gray-600 hover:bg-gray-100'
          }`}
        >
          Past Products
        </button>
      </div>

      {/* Main Content */}
      <main className="px-6 py-8">
        <div className="max-w-2xl mx-auto">
          {/* Message Display */}
          {message && (
            <div className={`mb-4 p-3 rounded-lg text-sm ${
              message.type === 'success' 
                ? 'bg-green-100 text-green-700 border border-green-200' 
                : message.type === 'error' 
                ? 'bg-red-100 text-red-700 border border-red-200'
                : 'bg-blue-100 text-blue-700 border border-blue-200'
            }`}>
              {message.text}
            </div>
          )}

          {selectedProduct ? (
            renderDetailsView()
          ) : (
            <>
              {currentView === 'live' ? (
                getLiveProducts().length > 0 ? (
                  renderProductList(getLiveProducts())
                ) : (
                  <div className="text-center py-12 bg-white rounded-lg">
                    <div className="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mx-auto mb-4">
                      <span className="text-gray-400 text-2xl">🏁</span>
                    </div>
                    <h3 className="text-gray-500 font-medium mb-2">No Live Products</h3>
                    <p className="text-gray-400 text-sm">
                      All products have completed their lifecycle
                    </p>
                  </div>
                )
              ) : (
                getPastProducts().length > 0 ? (
                  renderProductList(getPastProducts())
                ) : (
                  <div className="text-center py-12 bg-white rounded-lg">
                    <div className="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mx-auto mb-4">
                      <span className="text-gray-400 text-2xl">📦</span>
                    </div>
                    <h3 className="text-gray-500 font-medium mb-2">No Past Products</h3>
                    <p className="text-gray-400 text-sm">
                      No completed products found
                    </p>
                  </div>
                )
              )}
            </>
          )}
        </div>
      </main>
    </div>
  );
}