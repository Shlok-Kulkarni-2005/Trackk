"use client";
import React, { useState } from "react";
import { Menu, Edit, Bell, Filter, Plus } from "lucide-react";
import { useRouter } from "next/navigation";
import Sidebar from "../../components/sidebarm";
import { useAlertNotifications } from "../lib/useAlertNotifications";

export default function MobileInterface({ username }: { username: string }) {
  const { unreadCount: alertCount, isAuthenticated } = useAlertNotifications();
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const router = useRouter();

  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };

  const handleAddJob = (): void => {
    console.log("Add Job clicked");
    router.push("/workpanel");
  };

  const handleSeeAlerts = (): void => {
    console.log("See Alerts clicked");
    router.push("/sendalerts");
  };

  const handleProductList = (): void => {
    console.log("Product List clicked");
    router.push("/reports");
  };



  const handleUpdateDetails = (): void => {
    console.log("Update Details clicked");
    router.push("/dispatched");
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Sidebar Component */}
      <Sidebar
        isOpen={sidebarOpen}
        onClose={() => setSidebarOpen(false)}
        username={username}
      />
      {/* Header */}
      <header className="bg-white shadow-sm px-4 py-4 flex items-center justify-between sticky top-0 z-20">
        <button
          onClick={handleMenuClick}
          className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
          type="button"
          aria-label="Open menu"
        >
          <Menu className="w-6 h-6 text-blue-700" />
        </button>

        <h1 className="text-lg md:text-xl font-semibold text-blue-700 text-center flex-grow">
          Admin (Manager Login)
        </h1>

        <div className="w-8 h-8 md:w-10 md:h-10 bg-blue-600 rounded-full flex items-center justify-center">
          <span className="text-white font-medium text-base md:text-lg">M</span>
        </div>
      </header>

      {/* Main Content */}
      <main className="px-4 py-6 flex-grow">
        <div className="w-full max-w-md mx-auto space-y-6">
          {/* Top Row */}
          <div className="grid grid-cols-2 gap-4">
            {/* Work Panel */}
            <button
              onClick={handleAddJob}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-2 md:space-y-3 border border-gray-100"
              type="button"
              aria-label="Access work panel"
            >
              <div className="w-10 h-10 md:w-12 md:h-12 bg-blue-50 rounded-full flex items-center justify-center">
                <Edit className="w-5 h-5 md:w-6 md:h-6 text-blue-600" />
              </div>
              <span className="text-gray-700 font-medium text-xs md:text-sm text-center">
                Work Panel
              </span>
            </button>

            {/* Send Alerts */}
            <button
              onClick={handleSeeAlerts}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-2 md:space-y-3 border border-gray-100 relative"
              type="button"
              aria-label={`Send alerts (${alertCount} pending)`}
            >
              <div className="w-10 h-10 md:w-12 md:h-12 bg-blue-50 rounded-full flex items-center justify-center relative">
                <Bell className="w-5 h-5 md:w-6 md:h-6 text-blue-600" />
                {isAuthenticated && alertCount > 0 && (
                  <div className="absolute -top-1 -right-1 w-4 h-4 md:w-5 md:h-5 bg-red-500 rounded-full flex items-center justify-center">
                    <span className="text-white text-[10px] md:text-xs font-medium">
                      {alertCount}
                    </span>
                  </div>
                )}
              </div>
              <span className="text-gray-700 font-medium text-xs md:text-sm text-center">
                Send Alerts
              </span>
            </button>
          </div>

          {/* Middle Row - Reports and Dispatched Details */}
          <div className="grid grid-cols-2 gap-4">
            {/* Reports */}
            <button
              onClick={handleProductList}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-2 md:space-y-3 border border-gray-100"
              type="button"
              aria-label="View reports"
            >
              <div className="w-10 h-10 md:w-12 md:h-12 bg-blue-50 rounded-full flex items-center justify-center">
                <Filter className="w-5 h-5 md:w-6 md:h-6 text-blue-600" />
              </div>
              <span className="text-gray-700 font-medium text-xs md:text-sm text-center">
                Reports
              </span>
            </button>

            {/* Dispatched Details */}
            <button
              onClick={handleUpdateDetails}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-2 md:space-y-3 border border-gray-100"
              type="button"
              aria-label="View dispatched details"
            >
              <div className="w-10 h-10 md:w-12 md:h-12 bg-blue-50 rounded-full flex items-center justify-center">
                <Plus className="w-5 h-5 md:w-6 md:h-6 text-blue-600" />
              </div>
              <span className="text-gray-700 font-medium text-xs md:text-sm text-center">RFD Status</span>
            </button>
          </div>
        </div>
      </main>
    </div>
  );
}
