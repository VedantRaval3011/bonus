'use client';

import { useState, useMemo } from 'react';
import { ChevronLeft, ChevronRight, Search } from 'lucide-react';

interface Column {
  key: string;
  label: string;
  format?: (value: any) => string;
  render?: (value: any, row: any) => React.ReactNode; // Add this line
}

interface DataTableProps {
  data: any[];
  title: string;
  columns: Column[];
  searchTerm?: string;
  onSearchChange?: (term: string) => void;
  searchFields?: string[]; // Optional: specify fields to search
}

export default function DataTable({
  data,
  title,
  columns,
  searchTerm = '',
  onSearchChange,
  searchFields,
}: DataTableProps) {
  const [currentPage, setCurrentPage] = useState(1);
  const [localSearchTerm, setLocalSearchTerm] = useState(searchTerm);
  const itemsPerPage = 20;

  // Use local search term if onSearchChange is not provided
  const activeSearchTerm = onSearchChange ? searchTerm : localSearchTerm;

  // Filter data based on search term BEFORE pagination
  const filteredData = useMemo(() => {
    if (!data || !activeSearchTerm.trim()) return data || [];

    return data.filter((employee: any) => {
      const searchLower = activeSearchTerm.toLowerCase();
      // If searchFields are specified, search only those; otherwise, search all fields
      if (searchFields) {
        return searchFields.some((field) =>
          employee[field]?.toString().toLowerCase().includes(searchLower)
        );
      }
      return Object.values(employee).some(
        (value) =>
          value && value.toString().toLowerCase().includes(searchLower)
      );
    });
  }, [data, activeSearchTerm, searchFields]);

  // Reset to first page when search changes
  useMemo(() => {
    setCurrentPage(1);
  }, [activeSearchTerm]);

  // Handle empty data
  if (!data || data.length === 0) {
    return (
      <div className="bg-white p-4 rounded shadow-sm">
        <h3 className="text-lg font-semibold mb-2">{title}</h3>
        <p className="text-gray-500 text-sm">No data available</p>
      </div>
    );
  }

  // Pagination calculations
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const currentData = filteredData.slice(startIndex, startIndex + itemsPerPage);

  // Format cell values
  const formatValue = (value: any, formatter?: (value: any) => string) => {
    if (formatter) return formatter(value);
    if (typeof value === 'string' && value.includes('T') && value.includes('Z')) {
      return new Date(value).toLocaleDateString('en-IN');
    }
    return value?.toString() || '';
  };

  // Handle search input changes
  const handleSearchChange = (value: string) => {
    if (onSearchChange) {
      onSearchChange(value);
    } else {
      setLocalSearchTerm(value);
    }
  };

  return (
    <div className="bg-white rounded shadow-sm overflow-hidden">
      <div className="px-4 py-2 border-b bg-gray-50">
        <div className="flex justify-between items-center mb-2">
          <h3 className="text-lg font-semibold">{title}</h3>
          <span className="text-sm text-gray-600">
            {filteredData.length} of {data.length} employees
          </span>
        </div>

        {/* Local search input if not handled by parent */}
        {!onSearchChange && (
          <div className="relative flex items-center">
            <Search className="absolute left-2 top-1/2 transform -translate-y-1/2 h-4 w-4 text-gray-400" />
            <input
              type="text"
              placeholder="Search employees..."
              value={localSearchTerm}
              onChange={(e) => handleSearchChange(e.target.value)}
              className="w-full pl-8 pr-10 py-1.5 border rounded text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
            {localSearchTerm && (
              <button
                onClick={() => handleSearchChange('')}
                className="absolute right-2 top-1/2 transform -translate-y-1/2 text-gray-400 hover:text-gray-600"
                aria-label="Clear search"
              >
                ✕
              </button>
            )}
          </div>
        )}
      </div>

      {filteredData.length === 0 ? (
        <div className="p-4 text-center text-gray-500">
          <p>No employees match the search criteria</p>
          <p className="text-sm mt-1">Try adjusting your search terms</p>
        </div>
      ) : (
        <>
          <div className="overflow-x-auto">
            <table className="min-w-full text-sm">
              <thead className="bg-gray-50">
                <tr>
                  {columns.map((column, index) => (
                    <th
                      key={index}
                      className="px-3 py-2 text-left font-medium text-gray-700"
                    >
                      {column.label}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-200">
                {currentData.map((row, index) => (
                  <tr key={index} className="hover:bg-gray-50">
                    {columns.map((column, colIndex) => (
                      <td key={colIndex} className="px-3 py-2">
                        {formatValue(row[column.key], column.format)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {totalPages > 1 && (
            <div className="px-4 py-2 bg-gray-50 border-t flex justify-between items-center">
              <span className="text-sm text-gray-700">
                Showing {startIndex + 1} to{' '}
                {Math.min(startIndex + itemsPerPage, filteredData.length)} of{' '}
                {filteredData.length} results
                {activeSearchTerm && (
                  <span className="ml-1 text-blue-600">
                    (filtered from {data.length} total)
                  </span>
                )}
              </span>
              <div className="flex items-center space-x-2">
                <button
                  onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))}
                  disabled={currentPage === 1}
                  className="px-2 py-1 border rounded text-sm disabled:opacity-50 hover:bg-gray-100"
                >
                  <ChevronLeft className="h-4 w-4" />
                </button>
                <span className="text-sm">
                  Page {currentPage} of {totalPages}
                </span>
                <button
                  onClick={() =>
                    setCurrentPage((prev) => Math.min(prev + 1, totalPages))
                  }
                  disabled={currentPage === totalPages}
                  className="px-2 py-1 border rounded text-sm disabled:opacity-50 hover:bg-gray-100"
                >
                  <ChevronRight className="h-4 w-4" />
                </button>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}