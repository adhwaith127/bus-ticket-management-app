import { useState, useEffect, useRef } from 'react';
import { Route, Plus, Eye, Pencil, Search } from 'lucide-react';
import Modal from '../../components/Modal';
import { useFilteredList } from '../../assets/js/useFilteredList';
import api, { BASE_URL } from '../../assets/js/axiosConfig';
import { Button }   from '@/components/ui/button';
import { Badge }    from '@/components/ui/badge';
import { Input }    from '@/components/ui/input';
import { Skeleton } from '@/components/ui/skeleton';

const FARE_TYPES = [
  { value: '1', label: 'TABLE' },
  { value: '2', label: 'GRAPH' },
];

const ROUTE_FLAGS = [
  { name: 'half',       label: 'Half Fare'           },
  { name: 'luggage',    label: 'Luggage'              },
  { name: 'student',    label: 'Student Concession'   },
  { name: 'adjust',     label: 'Fare Adjustment'      },
  { name: 'conc',       label: 'General Concession'   },
  { name: 'ph',         label: 'PH Concession'        },
  { name: 'pass_allow', label: 'Pass Holders'         },
  { name: 'use_stop',   label: 'Stop-based Fare'      },
];

export default function RouteListing() {

  // ── Section 1: Listing state ─────────────────────────────────────────────
  const [routes, setRoutes]           = useState([]);
  const [busTypes, setBusTypes]       = useState([]);
  const [stages, setStages]           = useState([]);
  const [loading, setLoading]         = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalMode, setModalMode]     = useState('create');
  const [submitting, setSubmitting]   = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [showDeleted, setShowDeleted] = useState(false);

  const emptyForm = {
    route_code: '', route_name: '', min_fare: '', fare_type: '',
    bus_type: '', start_from: 0, is_deleted: false,
    half: false, luggage: false, student: false, adjust: false,
    conc: false, ph: false, pass_allow: false, use_stop: false,
    route_stages: [],
  };
  const [formData, setFormData] = useState(emptyForm);

  // ── Section 2: Wizard state ──────────────────────────────────────────────
  const [wizardStep, setWizardStep]         = useState(0); // 0=closed, 1,2,3
  const [wizardSubmitting, setWizardSubmitting] = useState(false);
  const stageNameRef = useRef(null);

  const emptyWizard = {
    route_code: '', route_name: '', no_of_stages: '3', min_fare: '', fare_type: '1',
    bus_type: '', half: false, luggage: false, student: false, adjust: false,
    conc: false, ph: false, pass_allow: false, use_stop: false,
    fare_list: [],
    fare_matrix: [],
    stages: [],
  };
  const [wizardData, setWizardData] = useState(emptyWizard);
  const [stageInput, setStageInput] = useState({ stage_name: '', distance: '' });

  // ── Section 2b: beforeunload guard during wizard ──────────────────────────
  useEffect(() => {
    if (wizardStep === 0) return;
    const handler = (e) => {
      e.preventDefault();
      e.returnValue = 'Route creation is in progress. Leaving now will discard your changes.';
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [wizardStep]);

  // ── Section 3: Search & filter ───────────────────────────────────────────
  const { filteredItems, searchTerm, setSearchTerm } = useFilteredList(
    routes,
    ['route_code', 'route_name']
  );

  // ── Section 4: Fetch on mount ────────────────────────────────────────────
  useEffect(() => { fetchBusTypes(); fetchStages(); }, []);
  useEffect(() => { fetchRoutes(); }, [showDeleted]);

  // ── Section 5: API calls ─────────────────────────────────────────────────
  const fetchRoutes = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/routes`, {
        params: { show_deleted: showDeleted }
      });
      setRoutes(res.data?.data || []);
    } catch (err) {
      console.error('Error fetching routes:', err);
      setRoutes([]);
    } finally {
      setLoading(false);
    }
  };

  const fetchBusTypes = async () => {
    try {
      const res = await api.get(`${BASE_URL}/masterdata/dropdowns/bus-types`);
      setBusTypes(res.data?.data || []);
    } catch (err) {
      console.error('Error fetching bus types:', err);
    }
  };

  const fetchStages = async () => {
    try {
      const res = await api.get(`${BASE_URL}/masterdata/dropdowns/stages`);
      setStages(res.data?.data || []);
    } catch (err) {
      console.error('Error fetching stages:', err);
    }
  };

  // ── Section 6: Edit modal submit ─────────────────────────────────────────
  const handleSubmit = async () => {
    setSubmitting(true);
    try {
      const response = await api.put(`${BASE_URL}/masterdata/routes/update/${editingItem.id}`, formData);
      if (response?.status === 200) {
        window.alert(response.data.message || 'Success');
        setIsModalOpen(false);
        setFormData(emptyForm);
        fetchRoutes();
      }
    } catch (err) {
      if (!err.response) return window.alert('Server unreachable. Try later.');
      const { data } = err.response;
      const firstError = data.errors ? Object.values(data.errors)[0][0] : data.message;
      window.alert(firstError || 'Validation failed');
    } finally {
      setSubmitting(false);
    }
  };

  // ── Section 7: Modal helpers ──────────────────────────────────────────────
  const openViewModal = (item) => {
    setFormData({ ...emptyForm, ...item, bus_type: item.bus_type, route_stages: item.route_stages || [] });
    setEditingItem(item);
    setModalMode('view');
    setIsModalOpen(true);
  };

  const openEditModal = (item) => {
    setFormData({ ...emptyForm, ...item, bus_type: item.bus_type, route_stages: item.route_stages || [] });
    setEditingItem(item);
    setModalMode('edit');
    setIsModalOpen(true);
  };

  const handleInputChange = (e) => {
    const { name, value, type, checked } = e.target;
    let processedValue = type === 'checkbox' ? checked : value;
    if (name === 'route_code') {
      processedValue = value.replace(/[^a-zA-Z0-9]/g, '').slice(0, 4);
    } else if (name === 'route_name') {
      processedValue = value.replace(/[^a-zA-Z0-9]/g, '').slice(0, 14);
    }
    setFormData(prev => ({ ...prev, [name]: processedValue }));
  };

  // ── Section 8: RouteStage helpers (edit modal) ───────────────────────────
  const addStage = () => {
    setFormData(prev => ({
      ...prev,
      route_stages: [...prev.route_stages, { stage: '', sequence_no: prev.route_stages.length + 1, distance: '', stage_local_lang: '' }]
    }));
  };

  const updateStage = (index, field, value) => {
    setFormData(prev => {
      const updated = [...prev.route_stages];
      updated[index] = { ...updated[index], [field]: value };
      return { ...prev, route_stages: updated };
    });
  };

  const removeStage = (index) => {
    setFormData(prev => {
      const updated = prev.route_stages.filter((_, i) => i !== index);
      updated.forEach((s, i) => s.sequence_no = i + 1);
      return { ...prev, route_stages: updated };
    });
  };

  const moveStage = (index, direction) => {
    if (direction === 'up' && index === 0) return;
    if (direction === 'down' && index === formData.route_stages.length - 1) return;
    setFormData(prev => {
      const updated = [...prev.route_stages];
      const targetIndex = direction === 'up' ? index - 1 : index + 1;
      [updated[index], updated[targetIndex]] = [updated[targetIndex], updated[index]];
      updated.forEach((s, i) => s.sequence_no = i + 1);
      return { ...prev, route_stages: updated };
    });
  };

  const isReadOnly    = modalMode === 'view';
  const getModalTitle = () => ({ view: 'Route Details', edit: 'Edit Route', create: 'Create Route' }[modalMode]);

  // ── Section 9: Wizard functions ──────────────────────────────────────────
  const openWizard = () => {
    setWizardData(emptyWizard);
    setStageInput({ stage_name: '', distance: '' });
    setWizardStep(1);
  };

  const closeWizard = (force = false) => {
    if (!force && wizardStep > 0) {
      if (!window.confirm('Cancel route creation? All unsaved data will be lost.')) return;
    }
    setWizardStep(0);
  };

  const handleWizardChange = (e) => {
    const { name, value, type, checked } = e.target;
    let processedValue = type === 'checkbox' ? checked : value;
    if (name === 'route_code') {
      processedValue = value.replace(/[^a-zA-Z0-9]/g, '').slice(0, 4);
    } else if (name === 'route_name') {
      processedValue = value.replace(/[^a-zA-Z0-9]/g, '').slice(0, 14);
    }
    setWizardData(prev => ({ ...prev, [name]: processedValue }));
  };

  const goToStep2 = () => {
    const { route_code, route_name, no_of_stages, min_fare, fare_type, bus_type } = wizardData;
    if (!route_code.trim() || !route_name.trim() || !no_of_stages || !min_fare || !fare_type || !bus_type) {
      window.alert('Please fill all required fields.');
      return;
    }
    const n = parseInt(no_of_stages);
    if (isNaN(n) || n < 2) {
      window.alert('Number of stages must be at least 2.');
      return;
    }
    if (parseInt(fare_type) === 2 && n <= 2) {
      window.alert('No of stages should be greater than 2 in Graph fare.');
      return;
    }
    const minFare = parseFloat(min_fare) || 0;
    const ft = parseInt(fare_type);
    if (ft === 1) {
      const fareList = Array(n).fill(0);
      setWizardData(prev => ({ ...prev, fare_list: fareList, fare_matrix: [] }));
    } else {
      // Lower-triangular: (n-1) rows, row i has (i+1) entries, pre-filled with min_fare
      const fareMatrix = Array.from({ length: n - 1 }, (_, i) => Array(i + 1).fill(minFare));
      setWizardData(prev => ({ ...prev, fare_list: [], fare_matrix: fareMatrix }));
    }
    setWizardStep(2);
  };

  const goToStep3 = () => {
    setWizardData(prev => ({ ...prev, stages: [] }));
    setStageInput({ stage_name: '', distance: '' });
    setWizardStep(3);
    setTimeout(() => stageNameRef.current?.focus(), 100);
  };

  const updateWizardFareList = (idx, value) => {
    setWizardData(prev => {
      const updated = [...prev.fare_list];
      updated[idx] = Number(value) || 0;
      return { ...prev, fare_list: updated };
    });
  };

  const updateWizardFareMatrix = (row, col, value) => {
    setWizardData(prev => {
      const updated = prev.fare_matrix.map((r, i) =>
        i === row ? r.map((c, j) => (j === col ? Number(value) || 0 : c)) : r
      );
      return { ...prev, fare_matrix: updated };
    });
  };

  const saveStageEntry = () => {
    if (!stageInput.stage_name.trim()) {
      window.alert('Stage name is required.');
      return;
    }
    setWizardData(prev => ({
      ...prev,
      stages: [...prev.stages, {
        stage_name: stageInput.stage_name.trim(),
        distance: stageInput.distance || '0',
      }]
    }));
    setStageInput({ stage_name: '', distance: '' });
    setTimeout(() => stageNameRef.current?.focus(), 50);
  };

  const submitWizard = async () => {
    setWizardSubmitting(true);
    try {
      const { fare_list, fare_matrix, stages, no_of_stages, ...routeInfo } = wizardData;
      const ft = parseInt(routeInfo.fare_type);
      const payload = {
        ...routeInfo,
        stages,
        ...(ft === 1 ? { fare_list } : { fare_matrix }),
      };
      const res = await api.post(`${BASE_URL}/masterdata/routes/create-wizard`, payload);
      if (res.status === 201) {
        window.alert(res.data.message || 'Route created successfully');
        closeWizard(true);
        fetchRoutes();
      }
    } catch (err) {
      if (!err.response) { window.alert('Server unreachable. Try later.'); return; }
      const { data } = err.response;
      window.alert(data.message || 'Failed to create route');
    } finally {
      setWizardSubmitting(false);
    }
  };

  const n = parseInt(wizardData.no_of_stages) || 0;
  const stagesEntered = wizardData.stages.length;
  const allStagesDone = stagesEntered === n && n > 0;
  const fareTypeLabel = wizardData.fare_type === '1' ? 'TABLE' : 'GRAPH';

  // ── Section 10: Render ────────────────────────────────────────────────────
  return (
    <div className="p-3 sm:p-5 lg:p-7 min-h-screen bg-slate-50">

      {/* ═══════════════════════════════════════════════════════════════════ */}
      {/* ROUTE CREATION WIZARD OVERLAY                                        */}
      {/* ═══════════════════════════════════════════════════════════════════ */}
      {wizardStep > 0 && (
        <div className="fixed inset-0 bg-slate-900/70 z-50 flex items-start justify-center overflow-y-auto py-6 px-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl">

            {/* Wizard header */}
            <div className="bg-slate-900 text-white px-6 py-4 rounded-t-2xl flex items-center justify-between">
              <div>
                <h2 className="text-lg font-bold tracking-wide">
                  {wizardStep === 1 && 'New Route — Step 1 of 3: Route Info'}
                  {wizardStep === 2 && 'Fare Entry — Step 2 of 3'}
                  {wizardStep === 3 && 'Stage Names — Step 3 of 3'}
                </h2>
              </div>
              {/* Step dots */}
              <div className="flex items-center gap-2">
                {[1, 2, 3].map(s => (
                  <div key={s} className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold border-2 ${
                    wizardStep === s ? 'bg-white text-slate-900 border-white' :
                    wizardStep > s  ? 'bg-slate-600 text-white border-slate-500' :
                                      'bg-transparent text-slate-400 border-slate-500'
                  }`}>{s}</div>
                ))}
              </div>
            </div>

            {/* Route info bar (steps 2 and 3) */}
            {wizardStep > 1 && (
              <div className="bg-slate-50 border-b border-slate-200 px-6 py-2 flex flex-wrap items-center gap-x-6 gap-y-1 text-sm text-slate-700">
                <span><strong>RouteCode:</strong> {wizardData.route_code}</span>
                <span><strong>Route Name:</strong> {wizardData.route_name}</span>
                <span><strong>MinFare:</strong> ₹{wizardData.min_fare}</span>
                <span className="font-semibold text-slate-700"><strong>FARE TYPE</strong> {fareTypeLabel}</span>
              </div>
            )}

            {/* ─── STEP 1: Route Info ───────────────────────────────────────── */}
            {wizardStep === 1 && (
              <div className="p-6 space-y-5">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-sm font-medium text-slate-700">Route Code *</label>
                    <input type="text" name="route_code" value={wizardData.route_code} onChange={handleWizardChange}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-500" />
                  </div>
                  <div className="space-y-1">
                    <label className="text-sm font-medium text-slate-700">Route Name *</label>
                    <input type="text" name="route_name" value={wizardData.route_name} onChange={handleWizardChange}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-500" />
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div className="space-y-1">
                    <label className="text-sm font-medium text-slate-700">No of Stages *</label>
                    <input type="number" name="no_of_stages" value={wizardData.no_of_stages} onChange={handleWizardChange} min="2"
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-500" />
                  </div>
                  <div className="space-y-1">
                    <label className="text-sm font-medium text-slate-700">Min Fare (₹) *</label>
                    <input type="number" name="min_fare" value={wizardData.min_fare} onChange={handleWizardChange} min="0"
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-500" />
                  </div>
                  <div className="space-y-1">
                    <label className="text-sm font-medium text-slate-700">Fare Type *</label>
                    <select name="fare_type" value={wizardData.fare_type} onChange={handleWizardChange}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-500 bg-white">
                      <option value="">-- Select --</option>
                      {FARE_TYPES.map(ft => <option key={ft.value} value={ft.value}>{ft.label}</option>)}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-sm font-medium text-slate-700">Bus Type *</label>
                    <select name="bus_type" value={wizardData.bus_type} onChange={handleWizardChange}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-500 bg-white">
                      <option value="">-- Select --</option>
                      {busTypes.map(bt => <option key={bt.id} value={bt.id}>{bt.name}</option>)}
                    </select>
                  </div>
                </div>

                {/* Allowables */}
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <p className="text-sm font-medium text-slate-700">Select Allowables</p>
                    <button type="button" onClick={() => {
                      const allOn = ROUTE_FLAGS.every(f => wizardData[f.name]);
                      setWizardData(prev => {
                        const upd = {};
                        ROUTE_FLAGS.forEach(f => upd[f.name] = !allOn);
                        return { ...prev, ...upd };
                      });
                    }} className="text-xs px-3 py-1 rounded-md border border-slate-300 text-slate-600 hover:bg-slate-50">
                      {ROUTE_FLAGS.every(f => wizardData[f.name]) ? 'Deselect All' : 'Select All'}
                    </button>
                  </div>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                    {ROUTE_FLAGS.map(flag => (
                      <label key={flag.name} className={`flex items-center gap-2 text-sm p-2 rounded-lg border cursor-pointer transition-colors ${wizardData[flag.name] ? 'bg-slate-800 text-white border-slate-800' : 'bg-white text-slate-600 border-slate-200'}`}>
                        <input type="checkbox" name={flag.name} checked={wizardData[flag.name] || false} onChange={handleWizardChange} className="sr-only" />
                        {flag.label}
                      </label>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {/* ─── STEP 2: Fare Entry ───────────────────────────────────────── */}
            {wizardStep === 2 && (
              <div className="p-6">
                {/* TABLE FARE */}
                {wizardData.fare_type === '1' && (
                  <div>
                    <p className="text-sm text-slate-600 mb-3">
                      Enter the fare amount for each number of stages traveled. Row 1 = 1 stage trip fare, Row 2 = 2 stages trip fare, etc.
                    </p>
                    <div className="border border-slate-200 rounded-xl overflow-hidden">
                      <table className="w-full">
                        <thead>
                          <tr className="bg-slate-800 text-white">
                            <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider">Stages Traveled</th>
                            <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider">Fare Amount (₹)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {wizardData.fare_list.map((fare, idx) => (
                            <tr key={idx} className={idx === 0 ? 'bg-slate-50' : 'hover:bg-slate-50'}>
                              <td className="px-4 py-3 text-sm font-medium text-slate-700">
                                {idx + 1} {idx === 0 ? 'Stage' : 'Stages'}
                                {idx === 0 && <span className="ml-2 text-xs text-slate-400">(locked)</span>}
                              </td>
                              <td className="px-4 py-3">
                                {idx === 0 ? (
                                  <input type="number" value={0} disabled
                                    className="w-40 px-3 py-1.5 border border-slate-200 rounded-lg bg-slate-100 text-slate-400 cursor-not-allowed text-sm" />
                                ) : (
                                  <input type="number" value={fare} min="0" onChange={e => updateWizardFareList(idx, e.target.value)}
                                    onBlur={e => {
                                      const minF = parseFloat(wizardData.min_fare) || 0;
                                      const val = parseFloat(e.target.value) || 0;
                                      if (val > 0 && val < minF) {
                                        window.alert(`Minimum Fare is ${minF}`);
                                      }
                                    }}
                                    className="w-40 px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-slate-500 focus:outline-none text-sm" />
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* GRAPH FARE — lower-triangular, matching EXE */}
                {wizardData.fare_type === '2' && (
                  <div>
                    <p className="text-sm text-slate-600 mb-3">
                      Enter the fare between each pair of stages. Each row = destination stage, each column = origin stage.
                    </p>
                    <div className="border border-slate-200 rounded-xl overflow-auto">
                      <table className="border-collapse">
                        <thead>
                          <tr className="bg-slate-800 text-white">
                            <th className="px-3 py-2 text-xs font-semibold sticky left-0 bg-slate-800 z-10 border-r border-slate-700">Stage</th>
                            {Array.from({ length: n - 1 }, (_, i) => (
                              <th key={i} className="px-3 py-2 text-xs font-semibold text-center min-w-[80px] border-r border-slate-700">Stg {i + 1}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {wizardData.fare_matrix.map((row, rIdx) => (
                            <tr key={rIdx} className="hover:bg-slate-50">
                              <td className="px-3 py-2 text-sm font-semibold text-slate-700 sticky left-0 bg-slate-100 border-r border-slate-300 z-10">Stg {rIdx + 2}</td>
                              {Array.from({ length: n - 1 }, (_, cIdx) => {
                                const isActive = cIdx <= rIdx;
                                return (
                                  <td key={cIdx} className={`px-2 py-2 text-center border-r border-slate-100 ${!isActive ? 'bg-slate-100' : ''}`}>
                                    {isActive ? (
                                      <input type="number" value={row[cIdx]} min="0"
                                        onChange={e => updateWizardFareMatrix(rIdx, cIdx, e.target.value)}
                                        onBlur={e => {
                                          const minF = parseFloat(wizardData.min_fare) || 0;
                                          const val = parseFloat(e.target.value) || 0;
                                          if (val > 0 && val < minF) {
                                            window.alert(`Minimum Fare is ${minF}`);
                                          }
                                        }}
                                        className="w-16 px-2 py-1 text-center border border-slate-300 bg-slate-50 rounded text-sm focus:ring-1 focus:ring-slate-500 focus:outline-none"
                                      />
                                    ) : (
                                      <span className="text-slate-300 text-xs">—</span>
                                    )}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ─── STEP 3: Stage Name Entries ───────────────────────────────── */}
            {wizardStep === 3 && (
              <div className="p-6">
                <div className="flex gap-6">
                  {/* LEFT: Entry form */}
                  <div className="w-72 flex-shrink-0 space-y-4">
                    <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
                      <div className="space-y-1">
                        <label className="text-sm font-medium text-slate-700">Stage Name</label>
                        <input
                          ref={stageNameRef}
                          type="text"
                          value={stageInput.stage_name}
                          onChange={e => {
                            const val = e.target.value.replace(/[^a-zA-Z0-9]/g, '').slice(0, 11);
                            setStageInput(prev => ({ ...prev, stage_name: val }));
                          }}
                          onKeyDown={e => { if (e.key === 'Enter') saveStageEntry(); }}
                          placeholder="Enter stage name"
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-500 bg-white"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-sm font-medium text-slate-700">Distance (km)</label>
                        <input
                          type="text"
                          inputMode="decimal"
                          value={stageInput.distance}
                          onChange={e => {
                            const val = e.target.value.replace(/[^0-9.]/g, '').slice(0, 11);
                            setStageInput(prev => ({ ...prev, distance: val }));
                          }}
                          onKeyDown={e => { if (e.key === 'Enter') saveStageEntry(); }}
                          placeholder="0"
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-500 bg-white"
                        />
                      </div>
                      <button
                        type="button"
                        onClick={saveStageEntry}
                        disabled={stagesEntered >= n}
                        className="w-full py-2 bg-slate-900 hover:bg-slate-700 text-white font-medium rounded-lg disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                      >
                        Save
                      </button>
                      <div className={`text-center text-sm font-semibold py-2 rounded-lg border-2 ${allStagesDone ? 'border-emerald-400 text-emerald-700 bg-emerald-50' : 'border-slate-300 text-slate-600 bg-white'}`}>
                        Entries: {stagesEntered}/{n}
                      </div>
                    </div>
                  </div>

                  {/* RIGHT: Entered stages list */}
                  <div className="flex-1">
                    <div className="border border-slate-200 rounded-xl overflow-hidden">
                      <table className="w-full">
                        <thead>
                          <tr className="bg-slate-800 text-white">
                            <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wider w-12">S.No</th>
                            <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wider">Stage Name</th>
                            <th className="px-3 py-2 text-left text-xs font-semibold uppercase tracking-wider w-20">KM</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {wizardData.stages.length === 0 ? (
                            <tr>
                              <td colSpan={3} className="px-4 py-8 text-center text-slate-400 text-sm">
                                No stages entered yet. Use the form on the left.
                              </td>
                            </tr>
                          ) : wizardData.stages.map((s, idx) => (
                            <tr key={idx} className="hover:bg-slate-50">
                              <td className="px-3 py-2 text-sm font-mono text-slate-600">{idx + 1}</td>
                              <td className="px-3 py-2 text-sm font-medium text-slate-800">{s.stage_name}</td>
                              <td className="px-3 py-2 text-sm text-slate-600">{s.distance}</td>
                            </tr>
                          ))}
                          {/* Remaining empty slots */}
                          {Array.from({ length: Math.max(0, n - stagesEntered) }, (_, i) => (
                            <tr key={`empty-${i}`} className="bg-slate-50/50">
                              <td className="px-3 py-2 text-sm text-slate-300">{stagesEntered + i + 1}</td>
                              <td className="px-3 py-2 text-sm text-slate-300 italic">— not entered —</td>
                              <td className="px-3 py-2"></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Wizard footer navigation */}
            <div className="px-6 py-4 border-t border-slate-100 flex items-center justify-between">
              <button
                type="button"
                onClick={() => {
                  if (wizardStep === 1) closeWizard();
                  else setWizardStep(s => s - 1);
                }}
                className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors"
              >
                {wizardStep === 1 ? 'Cancel' : '← Back'}
              </button>

              <div className="flex items-center gap-3">
                {wizardStep === 1 && (
                  <button type="button" onClick={goToStep2}
                    className="px-6 py-2 text-sm font-medium text-white bg-slate-900 hover:bg-slate-700 rounded-lg shadow transition-colors">
                    Next: Fare Entry →
                  </button>
                )}
                {wizardStep === 2 && (
                  <button type="button" onClick={goToStep3}
                    className="px-6 py-2 text-sm font-medium text-white bg-slate-900 hover:bg-slate-700 rounded-lg shadow transition-colors">
                    Next: Stage Names →
                  </button>
                )}
                {wizardStep === 3 && (
                  <button type="button" onClick={submitWizard} disabled={!allStagesDone || wizardSubmitting}
                    className="px-8 py-2 text-sm font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg shadow disabled:opacity-40 disabled:cursor-not-allowed transition-colors">
                    {wizardSubmitting ? 'Creating Route...' : 'Finish'}
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ═══════════════════════════════════════════════════════════════════ */}
      {/* MAIN LISTING                                                          */}
      {/* ═══════════════════════════════════════════════════════════════════ */}

      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-6 gap-3">
        <div className="flex items-center gap-3">
          <div className="p-2.5 rounded-xl bg-slate-900">
            <Route size={20} className="text-white" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-slate-800 tracking-tight">Routes</h1>
            <p className="text-slate-500 text-sm mt-0.5">Manage bus routes for your company</p>
          </div>
        </div>
        <Button onClick={openWizard} className="bg-slate-900 hover:bg-slate-700 text-white gap-2 shadow-sm">
          <Plus size={16} /> Create Route
        </Button>
      </div>

      {/* Stats bar */}
      <div className="flex flex-wrap gap-2 mb-5">
        <div className="flex items-center gap-1.5 bg-white border border-slate-200 rounded-lg px-3 py-1.5 text-sm shadow-xs">
          <span className="text-slate-500">Total</span>
          <span className="font-bold text-slate-800">{routes.length}</span>
        </div>
        <div className="flex items-center gap-1.5 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-1.5 text-sm">
          <span className="text-emerald-600">Active</span>
          <span className="font-bold text-emerald-700">{routes.filter(r => !r.is_deleted).length}</span>
        </div>
        <label className="flex items-center gap-2 bg-white border border-slate-200 rounded-lg px-3 py-1.5 text-sm cursor-pointer hover:bg-slate-50">
          <input type="checkbox" checked={showDeleted} onChange={() => setShowDeleted(p => !p)} className="w-3.5 h-3.5 rounded border-slate-300 accent-slate-900" />
          <span className="text-slate-600">Show deleted</span>
        </label>
      </div>

      {/* Table card */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">

        {/* Search */}
        <div className="px-4 py-3 border-b border-slate-100 flex items-center gap-2">
          <Search size={15} className="text-slate-400 shrink-0" />
          <Input
            placeholder="Search by code or name..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="border-0 shadow-none focus-visible:ring-0 text-sm h-8 px-0"
          />
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200">
                {['ID', 'Code', 'Name', 'Bus Type', 'Stops', 'Min Fare', 'Fare Type', 'Status', ''].map(h => (
                  <th key={h} className="px-5 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wider">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                Array.from({ length: 5 }).map((_, i) => (
                  <tr key={i}>
                    {[50, 80, 140, 100, 60, 80, 70, 70, 60].map((w, j) => (
                      <td key={j} className="px-5 py-3"><Skeleton className="h-4 rounded" style={{ width: w }} /></td>
                    ))}
                  </tr>
                ))
              ) : filteredItems.length === 0 ? (
                <tr><td colSpan="9" className="px-5 py-10 text-center text-slate-400 text-sm">No routes found.</td></tr>
              ) : filteredItems.map(item => (
                <tr key={item.id} className="hover:bg-slate-50/60 transition-colors">
                  <td className="px-5 py-3.5"><span className="font-mono text-slate-500 text-xs font-semibold">#{item.id}</span></td>
                  <td className="px-5 py-3.5"><span className="font-semibold text-slate-800 text-base">{item.route_code}</span></td>
                  <td className="px-5 py-3.5"><span className="text-slate-700 text-base">{item.route_name}</span></td>
                  <td className="px-5 py-3.5"><span className="text-slate-600 text-base">{item.bus_type_name || '—'}</span></td>
                  <td className="px-5 py-3.5"><span className="text-slate-600 text-base">{item.route_stages?.length || 0}</span></td>
                  <td className="px-5 py-3.5"><span className="text-slate-600 text-base">₹{item.min_fare}</span></td>
                  <td className="px-5 py-3.5">
                    {item.fare_type === 1
                      ? <Badge className="bg-emerald-100 text-emerald-700 border border-emerald-200 hover:bg-emerald-100 text-xs">TABLE</Badge>
                      : <Badge className="bg-slate-100 text-slate-700 border border-slate-200 hover:bg-slate-100 text-xs">GRAPH</Badge>
                    }
                  </td>
                  <td className="px-5 py-3.5">
                    {item.is_deleted
                      ? <Badge className="bg-red-100 text-red-700 border border-red-200 hover:bg-red-100">Deleted</Badge>
                      : <Badge className="bg-emerald-100 text-emerald-700 border border-emerald-200 hover:bg-emerald-100">Active</Badge>
                    }
                  </td>
                  <td className="px-5 py-3.5">
                    <div className="flex items-center justify-end gap-1.5">
                      <button onClick={() => openViewModal(item)} className="p-2 rounded-md bg-slate-900 text-white hover:bg-slate-700 transition-colors" title="View"><Eye size={16} /></button>
                      <button onClick={() => openEditModal(item)} className="p-2 rounded-md bg-slate-900 text-white hover:bg-slate-700 transition-colors" title="Edit"><Pencil size={16} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* ═══════════════════════════════════════════════════════════════════ */}
      {/* EDIT / VIEW MODAL                                                     */}
      {/* ═══════════════════════════════════════════════════════════════════ */}
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={getModalTitle()}>
        <div className="space-y-5 max-h-[70vh] overflow-y-auto">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Route Code *</label>
              <input type="text" name="route_code" value={formData.route_code} onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 read-only:bg-slate-50" />
            </div>
            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Route Name *</label>
              <input type="text" name="route_name" value={formData.route_name} onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 read-only:bg-slate-50" />
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Min Fare (₹) *</label>
              <input type="number" name="min_fare" value={formData.min_fare} onChange={handleInputChange} readOnly={isReadOnly} min="0"
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 read-only:bg-slate-50" />
            </div>
            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Fare Type *</label>
              {isReadOnly ? (
                <input type="text" readOnly value={FARE_TYPES.find(f => f.value === String(formData.fare_type))?.label || `Type ${formData.fare_type}`}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg bg-slate-50" />
              ) : (
                <select name="fare_type" value={formData.fare_type} onChange={handleInputChange}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 bg-white">
                  <option value="">-- Select Fare Type --</option>
                  {FARE_TYPES.map(ft => <option key={ft.value} value={ft.value}>{ft.label}</option>)}
                </select>
              )}
            </div>
            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Start From (Stage)</label>
              <input type="number" name="start_from" value={formData.start_from} onChange={handleInputChange} readOnly={isReadOnly} min="0"
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 read-only:bg-slate-50" />
            </div>
          </div>

          <div className="space-y-1">
            <label className="text-sm font-medium text-slate-700">Bus Type *</label>
            {isReadOnly ? (
              <input type="text" value={formData.bus_type_name || '—'} readOnly className="w-full px-3 py-2 border border-slate-300 rounded-lg bg-slate-50" />
            ) : (
              <select name="bus_type" value={formData.bus_type} onChange={handleInputChange}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 bg-white">
                <option value="">-- Select Bus Type --</option>
                {busTypes.map(bt => <option key={bt.id} value={bt.id}>{bt.name}</option>)}
              </select>
            )}
          </div>

          <div>
            <div className="flex items-center justify-between mb-3">
              <p className="text-sm font-medium text-slate-700">Allowed Options</p>
              {!isReadOnly && (
                <button type="button" onClick={() => {
                  const allSelected = ROUTE_FLAGS.every(f => formData[f.name]);
                  setFormData(prev => {
                    const updates = {};
                    ROUTE_FLAGS.forEach(f => updates[f.name] = !allSelected);
                    return { ...prev, ...updates };
                  });
                }} className="text-xs px-3 py-1 rounded-md border border-slate-300 text-slate-600 hover:bg-slate-50 transition-colors">
                  {ROUTE_FLAGS.every(f => formData[f.name]) ? 'Deselect All' : 'Select All'}
                </button>
              )}
            </div>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {ROUTE_FLAGS.map(flag => (
                <label key={flag.name} className={`flex items-center gap-2 text-sm p-2 rounded-lg border transition-colors ${formData[flag.name] ? 'bg-slate-800 text-white border-slate-800' : 'bg-white text-slate-600 border-slate-200'} ${isReadOnly ? 'cursor-default' : 'cursor-pointer'}`}>
                  <input type="checkbox" name={flag.name} checked={formData[flag.name] || false} onChange={handleInputChange} disabled={isReadOnly} className="sr-only" />
                  {flag.label}
                </label>
              ))}
            </div>
          </div>

          <div className="border-t border-slate-200 pt-5">
            <div className="flex items-center justify-between mb-3">
              <div>
                <h3 className="text-sm font-semibold text-slate-700">Route Stops</h3>
                <p className="text-xs text-slate-500">Sequence of stops on this route</p>
              </div>
              {!isReadOnly && (
                <button type="button" onClick={addStage}
                  className="text-xs bg-blue-50 text-blue-600 hover:bg-blue-100 px-3 py-1.5 rounded-md transition-colors font-medium">
                  + Add Stop
                </button>
              )}
            </div>
            {formData.route_stages.length === 0 ? (
              <p className="text-sm text-slate-400 text-center py-4">No stops added yet</p>
            ) : (
              <div className="space-y-2">
                {formData.route_stages.map((stop, idx) => (
                  <div key={idx} className="flex items-center gap-2 p-2 bg-slate-50 rounded-lg border border-slate-200">
                    <div className="w-10 h-10 flex items-center justify-center bg-slate-800 text-white rounded-lg font-semibold text-sm">{stop.sequence_no}</div>
                    <div className="flex-1">
                      {isReadOnly ? (
                        <input type="text" value={stop.stage_name || '—'} readOnly className="w-full px-3 py-2 border border-slate-300 rounded-lg bg-white text-sm" />
                      ) : (
                        <select value={stop.stage} onChange={(e) => updateStage(idx, 'stage', e.target.value)}
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg bg-white text-sm">
                          <option value="">-- Select Stage --</option>
                          {stages.map(s => <option key={s.id} value={s.id}>{s.stage_name} ({s.stage_code})</option>)}
                        </select>
                      )}
                    </div>
                    <div className="w-28">
                      <input type="number" placeholder="Dist (km)" value={stop.distance} onChange={(e) => updateStage(idx, 'distance', e.target.value)}
                        readOnly={isReadOnly} step="0.1"
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm read-only:bg-white" />
                    </div>
                    {!isReadOnly && (
                      <div className="flex flex-col gap-1">
                        <button type="button" onClick={() => moveStage(idx, 'up')} disabled={idx === 0}
                          className="p-1 text-slate-500 hover:text-slate-700 disabled:opacity-30">
                          <i className="fas fa-chevron-up text-xs"></i>
                        </button>
                        <button type="button" onClick={() => moveStage(idx, 'down')} disabled={idx === formData.route_stages.length - 1}
                          className="p-1 text-slate-500 hover:text-slate-700 disabled:opacity-30">
                          <i className="fas fa-chevron-down text-xs"></i>
                        </button>
                      </div>
                    )}
                    {!isReadOnly && (
                      <button type="button" onClick={() => removeStage(idx)}
                        className="p-2 text-red-500 hover:text-red-700 hover:bg-red-50 rounded-md transition-colors">
                        <i className="fas fa-trash text-sm"></i>
                      </button>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>

          {modalMode === 'edit' && (
            <div className="flex items-center space-x-3 p-3 bg-red-50 rounded-lg border border-red-100">
              <input type="checkbox" name="is_deleted" id="route_is_deleted" checked={formData.is_deleted || false} onChange={handleInputChange} className="w-4 h-4 rounded border-slate-300" />
              <label htmlFor="route_is_deleted" className="text-sm font-medium text-red-700">Mark as deleted</label>
            </div>
          )}

          <div className="flex items-center justify-end space-x-3 pt-6 border-t border-slate-100 mt-6">
            <button type="button" onClick={() => setIsModalOpen(false)}
              className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50">
              {isReadOnly ? 'Close' : 'Cancel'}
            </button>
            {!isReadOnly && (
              <button type="button" onClick={handleSubmit} disabled={submitting}
                className="px-4 py-2 text-sm font-medium text-white bg-slate-800 rounded-lg hover:bg-slate-700 shadow-md disabled:opacity-50">
                {submitting ? 'Saving...' : 'Update'}
              </button>
            )}
          </div>
        </div>
      </Modal>

    </div>
  );
}
