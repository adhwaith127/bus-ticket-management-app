import { useState, useEffect } from 'react';
import {
  Settings, Smartphone, Ticket, BadgeDollarSign,
  Wifi, ToggleRight, Save, CheckCircle2,
} from 'lucide-react';
import api, { BASE_URL } from '../../assets/js/axiosConfig';

// ── Section 1: Reusable Field Components ─────────────────────────────────────
// These components avoid repeating the same label+input pattern throughout the form
// Each renders one labelled input row with consistent styling

const TextField = ({ label, name, value, onChange, placeholder = '', loading = false }) => (
  <div className="space-y-2">
    <label className="text-sm font-semibold text-slate-700">{label}</label>
    {loading ? (
      <div className="w-full h-10 bg-slate-100 rounded-xl animate-pulse"></div>
    ) : (
      <input
        type="text" name={name} value={value ?? ''} onChange={onChange}
        placeholder={placeholder}
        className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent text-sm transition-all"
      />
    )}
  </div>
);

const NumberField = ({ label, name, value, onChange, loading = false }) => (
  <div className="space-y-2">
    <label className="text-sm font-semibold text-slate-700">{label}</label>
    {loading ? (
      <div className="w-full h-10 bg-slate-100 rounded-xl animate-pulse"></div>
    ) : (
      <input
        type="number" name={name} value={value ?? 0} onChange={onChange}
        className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent text-sm transition-all"
      />
    )}
  </div>
);

const Toggle = ({ label, name, value, onChange, loading = false }) => (
  <label className={`flex items-center justify-between p-3 rounded-xl border-2 transition-all duration-150 ${
    loading 
      ? 'border-slate-200 bg-slate-50 animate-pulse cursor-not-allowed' 
      : 'border-slate-200 cursor-pointer hover:border-slate-300 hover:bg-slate-50'
  }`}>
    <span className="text-sm text-slate-700 font-medium">{label}</span>
    {loading ? (
      <div className="w-10 h-5 bg-slate-200 rounded-full"></div>
    ) : (
      <div className={`relative w-10 h-5 rounded-full transition-colors ${value ? 'bg-slate-800' : 'bg-slate-300'}`}>
        <input type="checkbox" name={name} checked={value || false} onChange={onChange} className="sr-only" />
        <span className={`absolute top-0.5 left-0.5 w-4 h-4 bg-white rounded-full shadow-sm transition-transform ${value ? 'translate-x-5' : ''}`} />
      </div>
    )}
  </label>
);

// Section wrapper with consistent card styling and optional loading skeleton
const Section = ({ title, icon: Icon, children, loading = false }) => (
  <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
    <h2 className="text-base font-bold text-slate-800 mb-5 pb-3 border-b border-slate-100 flex items-center gap-2">
      {loading
        ? <div className="w-4 h-4 border-2 border-slate-300 border-t-slate-800 rounded-full animate-spin shrink-0"></div>
        : Icon && <Icon size={16} className="text-slate-600 shrink-0" />
      }
      {title}
    </h2>
    {children}
  </div>
);

export default function SettingsPage() {

  // ── Section 2: State Management ──────────────────────────────────────────────
  const [formData, setFormData] = useState({});   // Start with empty object for skeleton display
  const [loading, setLoading]   = useState(true);
  const [saving, setSaving]     = useState(false);
  const [saved, setSaved]       = useState(false);

  // ── Section 3: Data Fetching ─────────────────────────────────────────────────
  // Fetch settings on component mount
  useEffect(() => { fetchSettings(); }, []);

  const fetchSettings = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/settings`);
      // Populate form with actual data from server
      setFormData(res.data?.data || {});
    } catch (err) {
      console.error('Error fetching settings:', err);
      setFormData({});
    } finally {
      setLoading(false);
    }
  };

  const handleSave = async () => {
    setSaving(true);
    setSaved(false);
    try {
      const res = await api.put(`${BASE_URL}/masterdata/settings`, formData);
      if (res?.status === 200) {
        setFormData(res.data?.data || formData);
        setSaved(true);
        // Auto-hide success message after 2 seconds
        setTimeout(() => setSaved(false), 2000);
      }
    } catch (err) {
      if (!err.response) return window.alert('Server unreachable. Try later.');
      const { data } = err.response;
      const firstError = data.errors ? Object.values(data.errors)[0][0] : data.message;
      window.alert(firstError || 'Save failed');
    } finally {
      setSaving(false);
    }
  };

  // ── Section 4: Input Handlers ────────────────────────────────────────────────
  const handleChange = (e) => {
    const { name, value, type, checked } = e.target;
    setFormData(prev => ({ ...prev, [name]: type === 'checkbox' ? checked : value }));
  };

  // ── Section 5: Render ────────────────────────────────────────────────────────
  const SaveButton = ({ bottom = false }) => (
    <button
      type="button" onClick={handleSave} disabled={saving || loading}
      className={`flex items-center justify-center gap-2 px-5 py-2.5 text-sm font-medium text-white bg-slate-900 hover:bg-slate-700 rounded-xl shadow-sm disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${bottom ? 'w-full sm:w-auto' : ''}`}
    >
      {saving ? (
        <>
          <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
          Saving...
        </>
      ) : saved ? (
        <>
          <CheckCircle2 size={16} />
          Saved!
        </>
      ) : (
        <>
          <Save size={16} />
          Save Settings
        </>
      )}
    </button>
  );

  return (
    <div className="p-3 sm:p-5 lg:p-7 min-h-screen bg-slate-50">

      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-6 gap-3">
        <div className="flex items-center gap-3">
          <div className="p-2.5 rounded-xl bg-slate-900">
            <Settings size={20} className="text-white" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-slate-800 tracking-tight">Company Settings</h1>
            <p className="text-slate-500 text-sm mt-0.5">Configure device, fare, and display settings</p>
          </div>
        </div>
        <SaveButton />
      </div>

      <div className="space-y-6">

        {/* ── Device & Access ──────────────────────────────────────────────────── */}
        <Section title="Device & Access" icon={Smartphone} loading={loading}>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <TextField label="Device ID (PalmtecID)" name="palmtec_id" value={formData.palmtec_id} onChange={handleChange} loading={loading} />
            <TextField label="User Password" name="user_pwd" value={formData.user_pwd} onChange={handleChange} loading={loading} />
            <TextField label="Master Password" name="master_pwd" value={formData.master_pwd} onChange={handleChange} loading={loading} />
            <TextField label="Currency" name="currency" value={formData.currency} onChange={handleChange} placeholder="e.g. INR" loading={loading} />
            <NumberField label="Language Option" name="language_option" value={formData.language_option} onChange={handleChange} loading={loading} />
            <NumberField label="Report Flag" name="report_flag" value={formData.report_flag} onChange={handleChange} loading={loading} />
            <NumberField label="Report Font" name="report_font" value={formData.report_font} onChange={handleChange} loading={loading} />
            <NumberField label="Default Stage" name="default_stage" value={formData.default_stage} onChange={handleChange} loading={loading} />
            <NumberField label="Stage Updation Msg" name="stage_updation_msg" value={formData.stage_updation_msg} onChange={handleChange} loading={loading} />
          </div>
        </Section>

        {/* ── Ticket Display (Headers & Footers) ───────────────────────────────── */}
        <Section title="Ticket Display (Headers & Footers)" icon={Ticket} loading={loading}>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <TextField label="Main Display Line 1" name="main_display" value={formData.main_display} onChange={handleChange} loading={loading} />
            <TextField label="Main Display Line 2" name="main_display2" value={formData.main_display2} onChange={handleChange} loading={loading} />
            <TextField label="Header Line 1" name="header1" value={formData.header1} onChange={handleChange} loading={loading} />
            <TextField label="Header Line 2" name="header2" value={formData.header2} onChange={handleChange} loading={loading} />
            <TextField label="Header Line 3" name="header3" value={formData.header3} onChange={handleChange} loading={loading} />
            <TextField label="Footer Line 1" name="footer1" value={formData.footer1} onChange={handleChange} loading={loading} />
            <TextField label="Footer Line 2" name="footer2" value={formData.footer2} onChange={handleChange} loading={loading} />
          </div>
        </Section>


        {/* ── Fare Percentages & Amounts ───────────────────────────────────────── */}
        <Section title="Fare Percentages & Amounts" icon={BadgeDollarSign} loading={loading}>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <NumberField label="Half Fare (%)" name="half_per" value={formData.half_per} onChange={handleChange} loading={loading} />
            <NumberField label="Concession (%)" name="con_per" value={formData.con_per} onChange={handleChange} loading={loading} />
            <NumberField label="PH Concession (%)" name="phy_per" value={formData.phy_per} onChange={handleChange} loading={loading} />
            <NumberField label="Student Max Amount" name="st_max_amt" value={formData.st_max_amt} onChange={handleChange} loading={loading} />
            <NumberField label="Student Min Concession" name="st_min_con" value={formData.st_min_con} onChange={handleChange} loading={loading} />
            <NumberField label="Rounding Amount" name="round_amt" value={formData.round_amt} onChange={handleChange} loading={loading} />
            <NumberField label="Luggage Unit Rate" name="luggage_unit_rate" value={formData.luggage_unit_rate} onChange={handleChange} loading={loading} />
            <NumberField label="ST Roundoff Amount" name="st_roundoff_amt" value={formData.st_roundoff_amt} onChange={handleChange} loading={loading} />
          </div>
        </Section>

        {/* ── Communication & FTP ───────────────────────────────────────────────── */}
        <Section title="Communication & FTP" icon={Wifi} loading={loading}>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <TextField label="Phone 2" name="ph_no2" value={formData.ph_no2} onChange={handleChange} loading={loading} />
            <TextField label="Phone 3" name="ph_no3" value={formData.ph_no3} onChange={handleChange} loading={loading} />
            <TextField label="Access Point" name="access_point" value={formData.access_point} onChange={handleChange} loading={loading} />
            <TextField label="Destination" name="dest_adds" value={formData.dest_adds} onChange={handleChange} loading={loading} />
            <TextField label="FTP Username" name="username" value={formData.username} onChange={handleChange} loading={loading} />
            <TextField label="FTP Password" name="password" value={formData.password} onChange={handleChange} loading={loading} />
            <TextField label="Upload Path" name="uploadpath" value={formData.uploadpath} onChange={handleChange} loading={loading} />
            <TextField label="Download Path" name="downloadpath" value={formData.downloadpath} onChange={handleChange} loading={loading} />
            <TextField label="HTTP URL" name="http_url" value={formData.http_url} onChange={handleChange} loading={loading} />
          </div>
        </Section>

        {/* ── Feature Toggles ───────────────────────────────────────────────────── */}
        <Section title="Feature Toggles" icon={ToggleRight} loading={loading}>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
            <Toggle label="Round Off" name="roundoff" value={formData.roundoff} onChange={handleChange} loading={loading} />
            <Toggle label="Round Up" name="round_up" value={formData.round_up} onChange={handleChange} loading={loading} />
            <Toggle label="Remove Ticket Flag" name="remove_ticket_flag" value={formData.remove_ticket_flag} onChange={handleChange} loading={loading} />
            <Toggle label="Stage Font Flag" name="stage_font_flag" value={formData.stage_font_flag} onChange={handleChange} loading={loading} />
            <Toggle label="Next Fare Flag" name="next_fare_flag" value={formData.next_fare_flag} onChange={handleChange} loading={loading} />
            <Toggle label="Odometer Entry" name="odometer_entry" value={formData.odometer_entry} onChange={handleChange} loading={loading} />
            <Toggle label="Ticket No Big Font" name="ticket_no_big_font" value={formData.ticket_no_big_font} onChange={handleChange} loading={loading} />
            <Toggle label="Crew Check" name="crew_check" value={formData.crew_check} onChange={handleChange} loading={loading} />
            <Toggle label="GPRS Enable" name="gprs_enable" value={formData.gprs_enable} onChange={handleChange} loading={loading} />
            <Toggle label="Trip Send" name="tripsend_enable" value={formData.tripsend_enable} onChange={handleChange} loading={loading} />
            <Toggle label="Schedule Send" name="schedulesend_enable" value={formData.schedulesend_enable} onChange={handleChange} loading={loading} />
            <Toggle label="Send Pending" name="sendpend" value={formData.sendpend} onChange={handleChange} loading={loading} />
            <Toggle label="Inspector Report" name="inspect_rpt" value={formData.inspect_rpt} onChange={handleChange} loading={loading} />
            <Toggle label="ST Roundoff" name="st_roundoff_enable" value={formData.st_roundoff_enable} onChange={handleChange} loading={loading} />
            <Toggle label="ST Fare Edit" name="st_fare_edit" value={formData.st_fare_edit} onChange={handleChange} loading={loading} />
            <Toggle label="Multiple Pass" name="multiple_pass" value={formData.multiple_pass} onChange={handleChange} loading={loading} />
            <Toggle label="Simple Report" name="simple_report" value={formData.simple_report} onChange={handleChange} loading={loading} />
            <Toggle label="Inspector SMS" name="inspector_sms" value={formData.inspector_sms} onChange={handleChange} loading={loading} />
            <Toggle label="Auto Shutdown" name="auto_shut_down" value={formData.auto_shut_down} onChange={handleChange} loading={loading} />
            <Toggle label="User Password Enable" name="userpswd_enable" value={formData.userpswd_enable} onChange={handleChange} loading={loading} />
          </div>
        </Section>

        {/* Bottom save button for convenience on long pages */}
        <div className="flex justify-end pt-2">
          <SaveButton bottom />
        </div>

      </div>
    </div>
  );
}