import './App.css'
import { useState } from 'react'
import ExcelJS from 'exceljs'
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  PieChart,
  Pie,
  Cell,
  ResponsiveContainer,
  Legend
} from 'recharts'

function App() {

  const [tickets, setTickets] = useState([
    { id: 'JIRA-101', module: 'ENLISTMENT', status: 'TO DO', assignedTo: 'Piyush' },
    { id: 'JIRA-102', module: 'FEE', status: 'DUSTED', assignedTo: 'Rahul' },
    { id: 'JIRA-103', module: 'SECTIONING', status: 'DEV IN PROGRESS', assignedTo: 'Sneha' },
    { id: 'JIRA-104', module: 'ENLISTMENT', status: 'TO DO', assignedTo: 'Piyush' },
    { id: 'JIRA-105', module: 'SECTIONING', status: 'TO DO', assignedTo: 'Amit' }
  ])

  const [selectedPerson, setSelectedPerson] = useState('All Members')
  const [selectedModule, setSelectedModule] = useState('All Modules')
  const [selectedCampus, setSelectedCampus] = useState('All Campuses')
  const [searchTicket, setSearchTicket] = useState('')
  const [showAnalytics, setShowAnalytics] = useState(false)

  // =========================
  // SAFE EXCEL IMPORT
  // =========================
  const handleFileUpload = async (e) => {

    const file = e.target.files[0]
    if (!file) return

    const workbook = new ExcelJS.Workbook()
    const buffer = await file.arrayBuffer()
    await workbook.xlsx.load(buffer)

    const sheet = workbook.worksheets[0]
    const parsed = []

    sheet.eachRow((row, i) => {
      if (i === 1) return

      const ticket = {
        id: row.getCell(2).value?.toString(),
        module: row.getCell(3).value?.toString() || 'SECTIONING',
        status: row.getCell(4).value?.toString() || 'TO DO',
        assignedTo: row.getCell(7).value?.toString() || 'Unassigned'
      }

      if (ticket.id) parsed.push(ticket)
    })

    setTickets(prev => [...prev, ...parsed])
  }

  // =========================
  // FILTERS
  // =========================
  const filteredTickets = tickets.filter((t) =>
    (selectedPerson === 'All Members' || t.assignedTo === selectedPerson) &&
    (selectedModule === 'All Modules' || t.module === selectedModule) &&
    t.id.toLowerCase().includes(searchTicket.toLowerCase())
  )

  // =========================
  // TRACKER
  // =========================
  const totalTickets = tickets.length
  const dustedTickets = tickets.filter(t => t.status === 'DUSTED').length
  const pendingTickets = tickets.filter(t => t.status !== 'DUSTED').length

  // =========================
  // ANALYTICS DATA
  // =========================
  const moduleData = [
    {
      module: 'ENLISTMENT',
      pending: tickets.filter(t => t.module === 'ENLISTMENT' && t.status !== 'DUSTED').length
    },
    {
      module: 'SECTIONING',
      pending: tickets.filter(t => t.module === 'SECTIONING' && t.status !== 'DUSTED').length
    },
    {
      module: 'FEE',
      pending: tickets.filter(t => t.module === 'FEE' && t.status !== 'DUSTED').length
    }
  ]

  const employeeData = [
    { name: 'Piyush', value: tickets.filter(t => t.assignedTo === 'Piyush').length },
    { name: 'Rahul', value: tickets.filter(t => t.assignedTo === 'Rahul').length },
    { name: 'Sneha', value: tickets.filter(t => t.assignedTo === 'Sneha').length },
    { name: 'Amit', value: tickets.filter(t => t.assignedTo === 'Amit').length }
  ]

  const COLORS = ['#2563eb', '#16a34a', '#f97316', '#dc2626']

  return (

    <div className="dashboard">

      {/* HEADER */}
      <div className="header">

        <div className="logo-section">
          <img src="/logo.png" className="logo" />
          <div>
            <h1 className="main-title">Enlistment Command Center</h1>
            <p className="sub-title">MasterSoft ERP Support Dashboard</p>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '10px' }}>

          <label className="btn">
            Import Excel
            <input type="file" hidden onChange={handleFileUpload} />
          </label>

          <button className="btn" onClick={() => setShowAnalytics(!showAnalytics)}>
            Analytics
          </button>

        </div>

      </div>

      {/* =========================
          TRACKER CARDS
      ========================= */}
      <div className="card-container">

        <div className="card blue-card">
          <p>Total Tickets</p>
          <h2>{totalTickets}</h2>
        </div>

        <div className="card green-card">
          <p>DUSTED Tickets</p>
          <h2>{dustedTickets}</h2>
        </div>

        <div className="card orange-card">
          <p>Pending Tickets</p>
          <h2>{pendingTickets}</h2>
        </div>

      </div>

      {/* FILTERS */}
      <div className="filter-section">

        <select value={selectedPerson} onChange={(e) => setSelectedPerson(e.target.value)}>
          <option>All Members</option>
          <option>Piyush</option>
          <option>Rahul</option>
          <option>Sneha</option>
          <option>Amit</option>
        </select>

        <select value={selectedModule} onChange={(e) => setSelectedModule(e.target.value)}>
          <option>All Modules</option>
          <option>SECTIONING</option>
          <option>ENLISTMENT</option>
          <option>FEE</option>
        </select>

        <select value={selectedCampus} onChange={(e) => setSelectedCampus(e.target.value)}>
          <option>All Campuses</option>
          <option>LPU M</option>
          <option>LPU C</option>
        </select>

        <input
          placeholder="Search Ticket"
          value={searchTicket}
          onChange={(e) => setSearchTicket(e.target.value)}
        />

      </div>

      {/* TABLE */}
      <table>
        <thead>
          <tr>
            <th>ID</th>
            <th>Module</th>
            <th>Status</th>
            <th>Assigned</th>
          </tr>
        </thead>

        <tbody>
          {filteredTickets.map((t, i) => (
            <tr key={i}>
              <td>{t.id}</td>
              <td>{t.module}</td>
              <td>{t.status}</td>
              <td>{t.assignedTo}</td>
            </tr>
          ))}
        </tbody>
      </table>

      {/* =========================
          ANALYTICS SECTION
      ========================= */}
      {showAnalytics && (
        <div className="analytics-panel">

          <h2>Analytics Dashboard</h2>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>

            {/* LINE CHART */}
            <div className="card">
              <h3>Module Pending Tickets</h3>
              <ResponsiveContainer width="100%" height={250}>
                <LineChart data={moduleData}>
                  <XAxis dataKey="module" />
                  <YAxis />
                  <Tooltip />
                  <Line dataKey="pending" stroke="#2563eb" />
                </LineChart>
              </ResponsiveContainer>
            </div>

            {/* PIE CHART */}
            <div className="card">
              <h3>Employee Workload</h3>
              <ResponsiveContainer width="100%" height={250}>
                <PieChart>
                  <Pie data={employeeData} dataKey="value" outerRadius={90} label>
                    {employeeData.map((_, i) => (
                      <Cell key={i} fill={COLORS[i]} />
                    ))}
                  </Pie>
                  <Tooltip />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </div>

          </div>

        </div>
      )}

    </div>
  )
}

export default App
