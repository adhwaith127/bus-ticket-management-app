import { useState, useEffect } from 'react';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function FareEditor() {

  // ── Section 1: State ────────────────────────────────────────────────────
  const [routes, setRoutes]           = useState([]);
  const [selectedRoute, setSelectedRoute] = useState(null);
  const [stages, setStages]           = useState([]);
  const [fareMatrix, setFareMatrix]   = useState([]);
  const [loading, setLoading]         = useState(false);
  const [saving, setSaving]           = useState(false);
  const [hasChanges, setHasChanges]   = useState(false);

  // ── Section 2: Fetch routes on mount ─────────────────────────────────────
  useEffect(() => {
    fetchRoutes();
  }, []);

  const fetchRoutes = async () => {
    try {
      const res = await api.get(`${BASE_URL}/masterdata/routes/`);
      setRoutes(res.data?.data || []);
    } catch (err) {
      console.error('Error fetching routes:', err);
    }
  };

  // ── Section 3: Load fare data when route is selected ─────────────────────
  const handleRouteSelect = async (routeId) => {
    if (!routeId) {
      setSelectedRoute(null);
      setStages([]);
      setFareMatrix([]);
      return;
    }

    setLoading(true);
    setHasChanges(false);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/fares/editor/${routeId}/`);
      const { route, stages: stageList, fare_matrix } = res.data.data;
      
      setSelectedRoute(route);
      setStages(stageList);
      setFareMatrix(fare_matrix);
    } catch (err) {
      console.error('Error loading fare data:', err);
      window.alert('Failed to load fare data. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  // ── Section 4: Update fare value in matrix ───────────────────────────────
  const updateFare = (rowIdx, colIdx, value) => {
    const updatedMatrix = fareMatrix.map((row, i) =>
      i === rowIdx
        ? row.map((fare, j) => (j === colIdx ? Number(value) || 0 : fare))
        : row
    );
    setFareMatrix(updatedMatrix);
    setHasChanges(true);
  };

  // ── Section 5: Save fare table ───────────────────────────────────────────
  const handleSave = async () => {
    if (!selectedRoute) return;

    setSaving(true);
    try {
      const res = await api.post(
        `${BASE_URL}/masterdata/fares/update/${selectedRoute.id}/`,
        { fare_matrix: fareMatrix }
      );
      
      window.alert(res.data.message || 'Fare table saved successfully!');
      setHasChanges(false);
    } catch (err) {
      console.error('Error saving fare table:', err);
      if (!err.response) return window.alert('Server unreachable. Try later.');
      const { data } = err.response;
      const firstError = data.errors ? Object.values(data.errors)[0][0] : data.message;
      window.alert(firstError || 'Save failed');
    } finally {
      setSaving(false);
    }
  };

  // ── Section 6: Auto-fill helpers ─────────────────────────────────────────
  const autoFillSymmetric = () => {
    // Copy upper triangle to lower triangle (make matrix symmetric)
    const updated = fareMatrix.map((row, i) =>
      row.map((fare, j) => {
        if (i > j) {
          return fareMatrix[j][i]; // Mirror from upper triangle
        }
        return fare;
      })
    );
    setFareMatrix(updated);
    setHasChanges(true);
  };

  const clearAllFares = () => {
    if (!window.confirm('Clear all fares? This will reset the entire table to zero.')) return;
    const cleared = fareMatrix.map(row => row.map(() => 0));
    setFareMatrix(cleared);
    setHasChanges(true);
  };

  // ── Section 7: Render ─────────────────────────────────────────────────────
  return (
    <div className="p-6 md:p-10 min-h-screen bg-slate-50 animate-fade-in">

      {/* Header */}
      <div className="mb-8">
        <h1 className="text-3xl font-bold text-slate-800 tracking-tight">Fare Editor</h1>
        <p className="text-slate-500 mt-1">Manage stage-to-stage fare matrix for routes</p>
      </div>

      {/* Route Selector */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 mb-6">
        <div className="flex flex-col md:flex-row md:items-end gap-4">
          <div className="flex-1">
            <label className="block text-sm font-medium text-slate-700 mb-2">
              Select Route
            </label>
            <select
              value={selectedRoute?.id || ''}
              onChange={(e) => handleRouteSelect(e.target.value)}
              className="w-full px-4 py-2.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 bg-white text-slate-800"
            >
              <option value="">-- Choose a route to edit fares --</option>
              {routes.map(r => (
                <option key={r.id} value={r.id}>
                  {r.route_code} - {r.route_name} ({r.route_stages?.length || 0} stops)
                </option>
              ))}
            </select>
          </div>

          {selectedRoute && (
            <div className="flex gap-2">
              <button
                onClick={autoFillSymmetric}
                className="px-4 py-2.5 text-sm font-medium text-blue-700 bg-blue-50 border border-blue-200 rounded-lg hover:bg-blue-100 transition-colors"
              >
                <i className="fas fa-sync-alt mr-2"></i>
                Mirror Fares
              </button>
              <button
                onClick={clearAllFares}
                className="px-4 py-2.5 text-sm font-medium text-red-700 bg-red-50 border border-red-200 rounded-lg hover:bg-red-100 transition-colors"
              >
                <i className="fas fa-eraser mr-2"></i>
                Clear All
              </button>
            </div>
          )}
        </div>

        {selectedRoute && (
          <div className="mt-4 p-3 bg-blue-50 rounded-lg border border-blue-200">
            <p className="text-sm text-blue-800">
              <strong>Route:</strong> {selectedRoute.route_code} - {selectedRoute.route_name} |{' '}
              <strong>Fare Type:</strong> {selectedRoute.fare_type} |{' '}
              <strong>Stops:</strong> {stages.length}
            </p>
          </div>
        )}
      </div>

      {/* Fare Matrix Table */}
      {loading ? (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-12 text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-4 border-slate-200 border-t-slate-800 mx-auto mb-4"></div>
          <p className="text-slate-500">Loading fare data...</p>
        </div>
      ) : !selectedRoute ? (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-12 text-center">
          <i className="fas fa-map-marked-alt text-6xl text-slate-300 mb-4"></i>
          <p className="text-slate-500 text-lg">Select a route to start editing fares</p>
        </div>
      ) : stages.length === 0 ? (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-12 text-center">
          <i className="fas fa-exclamation-triangle text-6xl text-amber-300 mb-4"></i>
          <p className="text-slate-700 text-lg font-medium">No stops defined for this route</p>
          <p className="text-slate-500 mt-2">Please add route stops before editing fares</p>
        </div>
      ) : (
        <>
          {/* Instructions */}
          <div className="bg-gradient-to-r from-blue-50 to-indigo-50 rounded-xl border border-blue-200 p-4 mb-6">
            <div className="flex items-start gap-3">
              <i className="fas fa-info-circle text-blue-600 mt-1"></i>
              <div className="flex-1 text-sm text-slate-700">
                <p className="font-medium text-slate-800 mb-1">How to use:</p>
                <ul className="list-disc list-inside space-y-1 text-slate-600">
                  <li><strong>Diagonal cells (gray):</strong> Same origin-destination, usually ₹0</li>
                  <li><strong>Upper triangle:</strong> Forward journey fares (Stage A → Stage B)</li>
                  <li><strong>Lower triangle:</strong> Return journey fares (Stage B → Stage A)</li>
                  <li><strong>Mirror Fares:</strong> Copies upper triangle to lower (same fare both ways)</li>
                </ul>
              </div>
            </div>
          </div>

          {/* Scrollable table container */}
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full border-collapse">
                <thead>
                  <tr className="bg-slate-800">
                    <th className="sticky left-0 z-20 bg-slate-800 px-4 py-3 text-left text-xs font-semibold text-white uppercase tracking-wider border-r border-slate-700">
                      From → To
                    </th>
                    {stages.map((stage, idx) => (
                      <th
                        key={idx}
                        className="px-4 py-3 text-center text-xs font-semibold text-white uppercase tracking-wider min-w-[100px] border-r border-slate-700"
                      >
                        <div className="font-mono text-slate-300 text-[10px]">{stage.stage_code}</div>
                        <div>{stage.stage_name}</div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {stages.map((rowStage, rowIdx) => (
                    <tr key={rowIdx} className="hover:bg-slate-50 transition-colors">
                      {/* Row header (stage name) */}
                      <td className="sticky left-0 z-10 bg-slate-100 px-4 py-3 font-medium text-sm text-slate-800 border-r border-slate-300">
                        <div className="font-mono text-slate-500 text-[10px]">{rowStage.stage_code}</div>
                        <div>{rowStage.stage_name}</div>
                      </td>

                      {/* Fare cells */}
                      {stages.map((colStage, colIdx) => {
                        const isDiagonal = rowIdx === colIdx;
                        const isUpperTriangle = colIdx > rowIdx;
                        
                        return (
                          <td
                            key={colIdx}
                            className={`px-2 py-2 text-center border-r border-slate-200 ${
                              isDiagonal ? 'bg-slate-100' : ''
                            }`}
                          >
                            <input
                              type="number"
                              value={fareMatrix[rowIdx]?.[colIdx] || 0}
                              onChange={(e) => updateFare(rowIdx, colIdx, e.target.value)}
                              min="0"
                              step="1"
                              className={`w-full px-2 py-1.5 text-center border rounded-md focus:ring-2 focus:ring-blue-500 focus:outline-none text-sm ${
                                isDiagonal
                                  ? 'bg-slate-200 text-slate-400 cursor-not-allowed'
                                  : isUpperTriangle
                                  ? 'border-blue-300 bg-blue-50 font-semibold'
                                  : 'border-slate-300 bg-white'
                              }`}
                              disabled={isDiagonal}
                            />
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Save button */}
          <div className="mt-6 flex items-center justify-between bg-white rounded-2xl shadow-sm border border-slate-200 p-4">
            <div className="flex items-center gap-2">
              {hasChanges && (
                <span className="px-3 py-1 bg-amber-100 text-amber-800 text-xs font-medium rounded-full">
                  <i className="fas fa-exclamation-circle mr-1"></i>
                  Unsaved changes
                </span>
              )}
            </div>
            
            <button
              onClick={handleSave}
              disabled={saving || !hasChanges}
              className="px-6 py-2.5 bg-slate-800 hover:bg-slate-700 text-white font-medium rounded-lg shadow-md disabled:opacity-50 disabled:cursor-not-allowed transition-all"
            >
              {saving ? (
                <>
                  <i className="fas fa-spinner fa-spin mr-2"></i>
                  Saving...
                </>
              ) : (
                <>
                  <i className="fas fa-save mr-2"></i>
                  Save Fare Table
                </>
              )}
            </button>
          </div>
        </>
      )}

    </div>
  );
}
