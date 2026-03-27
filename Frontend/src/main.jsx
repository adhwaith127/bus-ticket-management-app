import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { createBrowserRouter, RouterProvider, Navigate } from 'react-router-dom'

import './index.css'

import Login from './pages/auth/Login'
import Signup from './pages/auth/Signup'
import Dashboard from './layouts/Dashboard'
import RoleBasedHome from './components/RoleBasedHome'
import ProtectedRoute from './components/ProtectedRoute'

import CompanyListing from './pages/listings/CompanyListing'
import UserListing from './pages/listings/UserListing'
import DepotListing from './pages/listings/DepotListing'
import BusTypeListing from './pages/listings/BusTypeListing'
import EmployeeTypeListing from './pages/listings/EmployeeTypeListing'
import CurrencyListing from './pages/listings/CurrencyListing'
import EmployeeListing from './pages/listings/EmployeeListing'
import VehicleListing from './pages/listings/VehicleListing'
import RouteListing from './pages/listings/RouteListing'

import CrewAssignmentListing from './pages/operations/CrewAssignmentListing'
import DealerManagement from './pages/operations/DealerManagement'
import DeviceApprovals from './pages/operations/DeviceApprovals'
import FareEditor from './pages/operations/FareEditor'

import TicketReport from './pages/reports/TicketReport'
import TripcloseReport from './pages/reports/TripcloseReport'
import SettlementPage from './pages/reports/SettlementPage'

import MdbImport from './pages/tools/MdbImport'
import SettingsPage from './pages/tools/SettingsPage'

import NotFound from './components/NotFound'

const router = createBrowserRouter([
  {
    path:'*',
    element:<NotFound />
  },
  {
    path: '/',
    element: <Navigate to="/login" replace />
  },
  {
    path: '/signup',
    element: <Signup />
  },
  {
    path: '/login',
    element: <Login />
  },
  {
    element: <ProtectedRoute />,
    children: [
      {
        path: '/dashboard',
        element: <Dashboard />,
        children: [
          {
            index: true,
            element: <RoleBasedHome />
          },
          {
            path: 'companies',
            element: <CompanyListing />
          },
          {
            path: 'users',
            element: <UserListing />
          },
          {
            path: 'depots',
            element: <DepotListing />
          },
          {
            path: 'ticket-report',
            element: <TicketReport />
          },
          {
            path: 'trip-close-report',
            element: <TripcloseReport />
          },
          {
            path: 'settlements',
            element: <SettlementPage />
          },
          {
            path: 'dealers',
            element: <DealerManagement />
          },
          {
            path: 'device-approvals',
            element: <DeviceApprovals />
          },
          {
            path: 'data-import',
            element: <MdbImport/>
          },
          {
            path: 'master-data/bus-types',
            element: <BusTypeListing />
          },
          {
            path: 'master-data/employee-types',
            element: <EmployeeTypeListing />
          },
          {
            path: 'master-data/currencies',
            element: <CurrencyListing />
          },
          {
            path: 'master-data/employees',
            element: <EmployeeListing />
          },
          {
            path: 'master-data/vehicles',
            element: <VehicleListing />
          },
          {
            path: 'master-data/routes',
            element: <RouteListing />
          },
          {
            path: 'master-data/fares',
            element: <FareEditor />
          },
          {
            path: 'master-data/crew-assignments',
            element: <CrewAssignmentListing />
          },
          {
            path: 'master-data/settings',
            element: <SettingsPage />
          },
        ]
      }
    ]
  }
]);

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <RouterProvider router={router} />
  </StrictMode>
);