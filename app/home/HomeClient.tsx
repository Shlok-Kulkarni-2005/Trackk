"use client";
import React, { useState } from "react";
import { Menu, Edit, Bell, Filter, Plus } from "lucide-react";
import { useRouter } from "next/navigation";
import Sidebar from "../../components/sidebar";
import { useAlertNotifications } from "../lib/useAlertNotifications";

export default function HomeClient({
  profileImage,
  username,
}: {
  profileImage: string | null;
  username: string | null;
}) {
  console.log("Sidebar username prop:", username); // Debug log
  const { unreadCount: alertCount, isAuthenticated } = useAlertNotifications();
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const router = useRouter();

  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };

  const handleAddJob = (): void => {
    router.push("/addjobs");
  };

  const handleSeeAlerts = (): void => {
    router.push("/alerts");
    console.log("See Alerts clicked");
  };

  const handleProductList = (): void => {
    router.push("/productlist");
    console.log("Product List clicked");
  };

  const handleProductCount = (): void => {
    router.push("/procount");
    console.log("Product Count clicked");
  };

  const handleUpdateDetails = (): void => {
    router.push("/updatedetailsop");
    console.log("Update Details clicked");
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
          Operator Panel
        </h1>

        {profileImage ? (
          <img
            src={profileImage}
            alt="Profile"
            className="w-8 h-8 md:w-10 md:h-10 rounded-full object-cover"
          />
        ) : (
          <div className="w-8 h-8 md:w-10 md:h-10 bg-blue-600 rounded-full flex items-center justify-center">
            <span className="text-white font-medium text-base md:text-lg">
              A
            </span>
          </div>
        )}
      </header>

      {/* Main Content */}
      <main className="px-4 py-6 flex-grow">
        <div className="w-full max-w-md mx-auto space-y-6">
          {/* Top Row */}
          <div className="grid grid-cols-2 gap-4">
            {/* Add Job */}
            <button
              onClick={handleAddJob}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-2 md:space-y-3 border border-gray-100"
              type="button"
              aria-label="Add new job"
            >
              <div className="w-10 h-10 md:w-12 md:h-12 bg-blue-50 rounded-full flex items-center justify-center">
                <Edit className="w-5 h-5 md:w-6 md:h-6 text-blue-600" />
              </div>
              <span className="text-gray-700 font-medium text-xs md:text-sm text-center">
                Add Job
              </span>
            </button>

            {/* See Alerts */}
            <button
              onClick={handleSeeAlerts}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-2 md:space-y-3 border border-gray-100 relative"
              type="button"
              aria-label={`View alerts (${alertCount} new)`}
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
                See Alerts!
              </span>
            </button>
          </div>

          {/* Middle Row */}
          <div className="grid grid-cols-2 gap-4">
            {/* Product List */}
            <button
              onClick={handleProductList}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-2 md:space-y-3 border border-gray-100"
              type="button"
              aria-label="View product list"
            >
              <div className="w-10 h-10 md:w-12 md:h-12 bg-blue-50 rounded-full flex items-center justify-center">
                <Filter className="w-5 h-5 md:w-6 md:h-6 text-blue-600" />
              </div>
              <span className="text-gray-700 font-medium text-xs md:text-sm text-center">
                Product List
              </span>
            </button>

            {/* Product Count */}
            <button
              onClick={handleProductCount}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-2 md:space-y-3 border border-gray-100"
              type="button"
              aria-label="View product count"
            >
              <div className="w-10 h-10 md:w-12 md:h-12 bg-blue-50 rounded-full flex items-center justify-center">
                <Plus className="w-5 h-5 md:w-6 md:h-6 text-blue-600" />
              </div>
              <span className="text-gray-700 font-medium text-xs md:text-sm text-center">
                Product Count
              </span>
            </button>
          </div>

          {/* Bottom Row */}
          <div className="flex justify-center">
            <button
              onClick={handleUpdateDetails}
              className="bg-white rounded-xl p-4 md:p-6 shadow-sm hover:shadow-md transition-all duration-200 flex flex-col items-center space-y-3 border border-gray-100 w-full max-w-[10rem]"
              type="button"
              aria-label="Update details"
            >
              <div className="w-10 h-10 md:w-12 md:h-12 bg-blue-50 rounded-full flex items-center justify-center">
                <Plus className="w-5 h-5 md:w-6 md:h-6 text-blue-600" />
              </div>
              <span className="text-gray-700 font-medium text-xs md:text-sm text-center">
                Update Details
              </span>
            </button>
          </div>
        </div>
      </main>
    </div>
  );
}
