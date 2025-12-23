import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { createBrowserRouter, RouterProvider, Navigate } from 'react-router-dom'

import './styles/base.css'

import Login from './pages/Login'
import Signup from './pages/Signup'
import Dashboard from './layouts/Dashboard'
import Home from './pages/Home'

import CompanyListing from './pages/CompanyListing'
import UserListing from './pages/UserListing'
import TicketReport from './pages/TicketReport'

const router = createBrowserRouter([
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
    path: '/dashboard',
    element: <Dashboard />,
    children: [
      {
        index: true,
        element: <Home />
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
        path:'ticket-report',
        element:<TicketReport/>
      }
    ]
  }
])

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <RouterProvider router={router} />
  </StrictMode>
)