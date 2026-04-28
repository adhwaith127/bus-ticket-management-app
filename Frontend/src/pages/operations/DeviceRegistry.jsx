import { useEffect, useMemo, useState } from "react";
import api, { BASE_URL } from "../../assets/js/axiosConfig";
import TableSkeleton from "../../components/TableSkeleton";
import Modal from "../../components/Modal";

// ── Helpers ──────────────────────────────────────────────────────────────────

const fmtDate = (v) => {
  if (!v) return "—";
  try {
    return new Date(v).toLocaleDateString(undefined, {
      day: "numeric", month: "short", year: "numeric",
    });
  } catch { return v; }
};

const LICENCE_BADGE = {
  Pending:  "bg-yellow-100 text-yellow-700 border-yellow-300",
  Active:   "bg-green-100  text-green-700  border-green-300",
  Inactive: "bg-slate-100  text-slate-600  border-slate-300",
  Expired:  "bg-red-100    text-red-700    border-red-300",
};

function StatusBadge({ status }) {
  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold border ${LICENCE_BADGE[status] ?? "bg-slate-100 text-slate-600 border-slate-300"}`}>
      {status}
    </span>
  );
}

function SummaryCard({ label, value, sub }) {
  return (
    <div className="rounded-xl border border-slate-200 bg-white shadow-sm p-4 flex flex-col gap-1">
      <p className="text-xs text-slate-500 font-medium uppercase tracking-wide">{label}</p>
      <p className="text-2xl font-bold text-slate-800">{value}</p>
      {sub && <p className="text-xs text-slate-400">{sub}</p>}
    </div>
  );
}

// ── Main Component ────────────────────────────────────────────────────────────

export default function DeviceRegistry() {
  const user = JSON.parse(localStorage.getItem("user") || "{}");
  const role = user?.role;
  const canManage = role === "superadmin" || role === "executive";
  const isDealerAdmin = role === "dealer_admin";

  const [devices, setDevices]       = useState([]);
  const [companies, setCompanies]   = useState([]);
  const [summary, setSummary]       = useState(null);
  const [loading, setLoading]       = useState(true);
  const [actionLoading, setActionLoading] = useState(null); // device id being actioned

  // Filters
  const [filterStatus,  setFilterStatus]  = useState("");
  const [filterCompany, setFilterCompany] = useState("");
  const [filterType,    setFilterType]    = useState("");

  // Assign modal
  const [assignModal, setAssignModal]     = useState(false);
  const [assignTarget, setAssignTarget]   = useState(null); // device
  const [assignForm, setAssignForm]       = useState({ company: "", display_name: "" });
  const [assignSubmitting, setAssignSubmitting] = useState(false);
  const [assignError, setAssignError]     = useState("");

  // ── Fetch ──────────────────────────────────────────────────────────────────
  const fetchAll = async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (filterStatus)  params.set("status", filterStatus);
      if (filterCompany) params.set("company", filterCompany);

      const [devRes, compRes, sumRes] = await Promise.all([
        api.get(`${BASE_URL}/etm-devices?${params}`),
        api.get(`${BASE_URL}/customer-data`),
        api.get(`${BASE_URL}/etm-devices/summary`),
      ]);
      setDevices(devRes.data?.data ?? []);
      setCompanies(compRes.data?.data ?? []);
      setSummary(sumRes.data?.data ?? null);
    } catch (err) {
      console.error("DeviceRegistry fetch error:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { fetchAll(); }, [filterStatus, filterCompany]);

  // ── Actions ────────────────────────────────────────────────────────────────
  const handleApprove = async (device) => {
    if (!window.confirm(`Approve device ${device.serial_number}? This will call the license server.`)) return;
    setActionLoading(device.id);
    try {
      await api.post(`${BASE_URL}/etm-devices/${device.id}/approve`);
      fetchAll();
    } catch (err) {
      alert(err?.response?.data?.error || "Approval failed");
    } finally {
      setActionLoading(null);
    }
  };

  const handleRevoke = async (device) => {
    if (!window.confirm(`Revoke device ${device.serial_number}?`)) return;
    setActionLoading(device.id);
    try {
      await api.post(`${BASE_URL}/etm-devices/${device.id}/revoke`);
      fetchAll();
    } catch (err) {
      alert(err?.response?.data?.error || "Revoke failed");
    } finally {
      setActionLoading(null);
    }
  };

  const handleCheckStatus = async (device) => {
    setActionLoading(device.id);
    try {
      await api.post(`${BASE_URL}/etm-devices/${device.id}/check-status`);
      fetchAll();
    } catch (err) {
      alert(err?.response?.data?.error || "Status check failed");
    } finally {
      setActionLoading(null);
    }
  };

  const openAssign = (device) => {
    setAssignTarget(device);
    setAssignForm({ company: device.company ?? "", display_name: device.display_name ?? "" });
    setAssignError("");
    setAssignModal(true);
  };

  const handleAssignSubmit = async () => {
    if (!assignForm.company) { setAssignError("Select a company"); return; }
    setAssignSubmitting(true);
    setAssignError("");
    try {
      await api.post(`${BASE_URL}/etm-devices/${assignTarget.id}/assign`, assignForm);
      setAssignModal(false);
      fetchAll();
    } catch (err) {
      setAssignError(err?.response?.data?.error || "Assignment failed");
    } finally {
      setAssignSubmitting(false);
    }
  };

  // ── Filtered list ──────────────────────────────────────────────────────────
  const filtered = useMemo(() => {
    let list = devices;
    if (filterType) list = list.filter(d => d.device_type === filterType);
    return list;
  }, [devices, filterType]);

  // ── Summary cards ──────────────────────────────────────────────────────────
  const byStatus = summary?.by_status ?? {};

  return (
    <div className="p-6 md:p-10 min-h-screen bg-slate-50 space-y-6 animate-fade-in">
      {/* Header */}
      <div>
        <h1 className="text-xl font-bold text-slate-800">ETM Device Registry</h1>
        <p className="text-sm text-slate-500 mt-0.5">
          Physical ETM machines and Android devices registered in the system.
        </p>
      </div>

      {/* Summary cards */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <SummaryCard label="Total Devices"  value={summary?.total ?? "—"} />
        <SummaryCard label="Active"         value={byStatus.Active  ?? "—"} sub="Licensed & running" />
        <SummaryCard label="Pending"        value={byStatus.Pending ?? "—"} sub="Awaiting approval" />
        <SummaryCard label="Expired"        value={byStatus.Expired ?? "—"} sub="License expired" />
      </div>

      {/* Filters */}
      <div className="flex flex-wrap gap-2">
        <select
          value={filterStatus}
          onChange={e => setFilterStatus(e.target.value)}
          className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm text-slate-700 bg-white focus:outline-none focus:ring-2 focus:ring-slate-300"
        >
          <option value="">All statuses</option>
          <option value="Pending">Pending</option>
          <option value="Active">Active</option>
          <option value="Inactive">Inactive</option>
          <option value="Expired">Expired</option>
        </select>

        <select
          value={filterType}
          onChange={e => setFilterType(e.target.value)}
          className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm text-slate-700 bg-white focus:outline-none focus:ring-2 focus:ring-slate-300"
        >
          <option value="">All types</option>
          <option value="ETM">ETM</option>
          <option value="ANDROID">Android</option>
        </select>

        {!isDealerAdmin && (
          <select
            value={filterCompany}
            onChange={e => setFilterCompany(e.target.value)}
            className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm text-slate-700 bg-white focus:outline-none focus:ring-2 focus:ring-slate-300"
          >
            <option value="">All companies</option>
            {companies.map(c => (
              <option key={c.id} value={c.id}>{c.company_name}</option>
            ))}
          </select>
        )}

        <button
          onClick={fetchAll}
          className="ml-auto border border-slate-200 rounded-lg px-3 py-1.5 text-sm text-slate-600 bg-white hover:bg-slate-50 transition-colors"
        >
          Refresh
        </button>
      </div>

      {/* Table */}
      <div className="rounded-xl border border-slate-200 bg-white shadow-sm overflow-x-auto">
        <table className="w-full text-sm min-w-[780px]">
          <thead className="bg-slate-50">
            <tr>
              <th className="px-4 py-3 text-left font-semibold text-slate-700">Serial No.</th>
              <th className="px-4 py-3 text-left font-semibold text-slate-700">Type</th>
              <th className="px-4 py-3 text-left font-semibold text-slate-700">Company</th>
              {!isDealerAdmin && (
                <th className="px-4 py-3 text-left font-semibold text-slate-700">Dealer</th>
              )}
              <th className="px-4 py-3 text-left font-semibold text-slate-700">Status</th>
              <th className="px-4 py-3 text-left font-semibold text-slate-700">Licence Expiry</th>
              {canManage && (
                <th className="px-4 py-3 text-left font-semibold text-slate-700">Actions</th>
              )}
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <TableSkeleton
                columns={canManage
                  ? ["w-28","w-16","w-32","w-24","w-20","w-24","w-28"]
                  : ["w-28","w-16","w-32","w-20","w-24"]}
              />
            ) : filtered.length === 0 ? (
              <tr>
                <td
                  colSpan={canManage ? (isDealerAdmin ? 6 : 7) : (isDealerAdmin ? 5 : 6)}
                  className="px-4 py-8 text-center text-slate-400"
                >
                  No devices found
                </td>
              </tr>
            ) : (
              filtered.map(device => {
                const expiring = device.days_until_expiry !== null
                  && device.days_until_expiry >= 0
                  && device.days_until_expiry <= 10;
                const expired = device.is_expired;
                const isActioning = actionLoading === device.id;

                return (
                  <tr key={device.id} className="border-t border-slate-100 hover:bg-slate-50">
                    {/* Serial + name */}
                    <td className="px-4 py-3">
                      <p className="font-medium text-slate-800 font-mono text-xs">{device.serial_number}</p>
                      {device.display_name && device.display_name !== device.serial_number && (
                        <p className="text-slate-400 text-xs mt-0.5">{device.display_name}</p>
                      )}
                      {device.mac_address && (
                        <p className="text-slate-300 text-[10px] font-mono mt-0.5">{device.mac_address}</p>
                      )}
                    </td>

                    {/* Type */}
                    <td className="px-4 py-3">
                      <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold ${
                        device.device_type === "ETM"
                          ? "bg-blue-50 text-blue-700"
                          : "bg-purple-50 text-purple-700"
                      }`}>
                        {device.device_type}
                      </span>
                    </td>

                    {/* Company */}
                    <td className="px-4 py-3">
                      {device.company_name
                        ? <p className="text-slate-700 font-medium">{device.company_name}</p>
                        : <p className="text-slate-300 italic text-xs">Unassigned</p>
                      }
                    </td>

                    {/* Dealer (hidden for dealer_admin — they know their own dealer) */}
                    {!isDealerAdmin && (
                      <td className="px-4 py-3 text-slate-500 text-xs">
                        {device.dealer_name ?? "—"}
                      </td>
                    )}

                    {/* Status */}
                    <td className="px-4 py-3">
                      <StatusBadge status={device.licence_status} />
                    </td>

                    {/* Expiry */}
                    <td className="px-4 py-3">
                      {device.licence_active_to ? (
                        <div>
                          <p className={`text-xs font-medium ${expired ? "text-red-600" : expiring ? "text-orange-500" : "text-slate-600"}`}>
                            {fmtDate(device.licence_active_to)}
                          </p>
                          {device.days_until_expiry !== null && (
                            <p className={`text-[10px] mt-0.5 ${expired ? "text-red-400" : expiring ? "text-orange-400" : "text-slate-400"}`}>
                              {expired
                                ? `Expired ${Math.abs(device.days_until_expiry)}d ago`
                                : `${device.days_until_expiry}d remaining`}
                            </p>
                          )}
                        </div>
                      ) : (
                        <span className="text-slate-300 text-xs">—</span>
                      )}
                    </td>

                    {/* Actions — superadmin / executive only */}
                    {canManage && (
                      <td className="px-4 py-3">
                        <div className="flex flex-wrap gap-1.5">
                          {/* Assign — available if no company yet */}
                          {!device.company && (
                            <button
                              onClick={() => openAssign(device)}
                              disabled={isActioning}
                              className="px-2.5 py-1 rounded-lg text-xs font-medium bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100 transition-colors disabled:opacity-50"
                            >
                              Assign
                            </button>
                          )}

                          {/* Re-assign — change company after assigned */}
                          {device.company && device.licence_status !== "Active" && (
                            <button
                              onClick={() => openAssign(device)}
                              disabled={isActioning}
                              className="px-2.5 py-1 rounded-lg text-xs font-medium bg-slate-50 text-slate-600 border border-slate-200 hover:bg-slate-100 transition-colors disabled:opacity-50"
                            >
                              Reassign
                            </button>
                          )}

                          {/* Approve — only for Pending/Inactive devices with a company */}
                          {device.company && (device.licence_status === "Pending" || device.licence_status === "Inactive") && (
                            <button
                              onClick={() => handleApprove(device)}
                              disabled={isActioning}
                              className="px-2.5 py-1 rounded-lg text-xs font-medium bg-green-50 text-green-700 border border-green-200 hover:bg-green-100 transition-colors disabled:opacity-50"
                            >
                              {isActioning ? "..." : "Approve"}
                            </button>
                          )}

                          {/* Check Status — for Active devices */}
                          {device.licence_status === "Active" && (
                            <button
                              onClick={() => handleCheckStatus(device)}
                              disabled={isActioning}
                              className="px-2.5 py-1 rounded-lg text-xs font-medium bg-indigo-50 text-indigo-700 border border-indigo-200 hover:bg-indigo-100 transition-colors disabled:opacity-50"
                            >
                              {isActioning ? "..." : "Sync"}
                            </button>
                          )}

                          {/* Revoke — active devices */}
                          {device.licence_status === "Active" && (
                            <button
                              onClick={() => handleRevoke(device)}
                              disabled={isActioning}
                              className="px-2.5 py-1 rounded-lg text-xs font-medium bg-red-50 text-red-600 border border-red-200 hover:bg-red-100 transition-colors disabled:opacity-50"
                            >
                              Revoke
                            </button>
                          )}
                        </div>
                      </td>
                    )}
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Assign Modal */}
      <Modal isOpen={assignModal} onClose={() => setAssignModal(false)}>
        <div className="space-y-4 w-full max-w-md">
          <div>
            <h2 className="text-base font-bold text-slate-800">
              {assignTarget?.company ? "Reassign Device" : "Assign Device"}
            </h2>
            <p className="text-xs text-slate-400 mt-0.5 font-mono">{assignTarget?.serial_number}</p>
          </div>

          {assignError && (
            <p className="text-xs text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2">
              {assignError}
            </p>
          )}

          <div className="space-y-3">
            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">Company *</label>
              <select
                value={assignForm.company}
                onChange={e => setAssignForm(f => ({ ...f, company: e.target.value }))}
                className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-300"
              >
                <option value="">— Select company —</option>
                {companies.map(c => (
                  <option key={c.id} value={c.id}>{c.company_name}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">Display Name</label>
              <input
                type="text"
                value={assignForm.display_name}
                onChange={e => setAssignForm(f => ({ ...f, display_name: e.target.value }))}
                placeholder="e.g. Depot A - ETM 01"
                className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-300"
              />
            </div>
          </div>

          <div className="flex justify-end gap-2 pt-1">
            <button
              onClick={() => setAssignModal(false)}
              className="px-4 py-2 rounded-lg text-sm font-medium border border-slate-200 text-slate-600 hover:bg-slate-50 transition-colors"
            >
              Cancel
            </button>
            <button
              onClick={handleAssignSubmit}
              disabled={assignSubmitting}
              className="px-4 py-2 rounded-lg text-sm font-medium bg-slate-900 text-white hover:bg-slate-700 transition-colors disabled:opacity-50"
            >
              {assignSubmitting ? "Saving…" : "Assign"}
            </button>
          </div>
        </div>
      </Modal>
    </div>
  );
}
