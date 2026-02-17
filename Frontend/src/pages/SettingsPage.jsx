import { useState, useEffect } from 'react';
import api, { BASE_URL } from '../assets/js/axiosConfig';

// ── Reusable field components ─────────────────────────────────────────────────
// These small helpers avoid repeating the same label+input pattern
// 20+ times. Each renders one labelled input row.

const TextField = ({ label, name, value, onChange, placeholder = '' }) => (
  <div className="space-y-1">
    <label className="text-sm font-medium text-slate-700">{label}</label>
    <input
      type="text" name={name} value={value ?? ''} onChange={onChange}
      placeholder={placeholder}
      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-400 text-sm"
    />
  </div>
);

const NumberField = ({ label, name, value, onChange }) => (
  <div className="space-y-1">
    <label className="text-sm font-medium text-slate-700">{label}</label>
    <input
      type="number" name={name} value={value ?? 0} onChange={onChange}
      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-400 text-sm"
    />
  </div>
);

const Toggle = ({ label, name, value, onChange }) => (
  <label className="flex items-center justify-between p-3 rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors">
    <span className="text-sm text-slate-700">{label}</span>
    <div className={`relative w-10 h-5 rounded-full transition-colors ${value ? 'bg-slate-800' : 'bg-slate-300'}`}>
      <input type="checkbox" name={name} checked={value || false} onChange={onChange} className="sr-only" />
      <span className={`absolute top-0.5 left-0.5 w-4 h-4 bg-white rounded-full shadow transition-transform ${value ? 'translate-x-5' : ''}`} />
    </div>
  </label>
);

// Section wrapper — gives each group a title and a card box
const Section = ({ title, children }) => (
  <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
    <h2 className="text-base font-semibold text-slate-800 mb-4 pb-2 border-b border-slate-100">{title}</h2>
    {children}
  </div>
);

export default function SettingsPage() {

  // ── Section 1: State ────────────────────────────────────────────────────────
  const [formData, setFormData] = useState(null);   // null = not loaded yet
  const [loading, setLoading]   = useState(true);
  const [saving, setSaving]     = useState(false);
  const [saved, setSaved]       = useState(false);  // brief success indicator

  // ── Section 2: Fetch on mount ────────────────────────────────────────────────
  useEffect(() => { fetchSettings(); }, []);

  // ── Section 3: API calls ─────────────────────────────────────────────────────
  const fetchSettings = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/settings/`);
      // If the company has no settings yet, data will be null — we use an empty object
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
      const res = await api.put(`${BASE_URL}/masterdata/settings/`, formData);
      if (res?.status === 200) {
        setFormData(res.data?.data || formData);
        setSaved(true);
        // Reset the "Saved!" label after 2 seconds
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

  // ── Section 4: Input handlers ─────────────────────────────────────────────────
  const handleChange = (e) => {
    const { name, value, type, checked } = e.target;
    setFormData(prev => ({ ...prev, [name]: type === 'checkbox' ? checked : value }));
  };

  // ── Section 5: Render ─────────────────────────────────────────────────────────
  if (loading) {
    return (
      <div className="p-6 md:p-10 min-h-screen bg-slate-50 flex items-center justify-center">
        <p className="text-slate-500">Loading settings...</p>
      </div>
    );
  }

  return (
    <div className="p-6 md:p-10 min-h-screen bg-slate-50 animate-fade-in">

      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
        <div>
          <h1 className="text-3xl font-bold text-slate-800 tracking-tight">Company Settings</h1>
          <p className="text-slate-500 mt-1">Configure device, fare, and display settings</p>
        </div>
        <button
          type="button" onClick={handleSave} disabled={saving}
          className="px-6 py-2.5 text-sm font-medium text-white bg-slate-800 rounded-xl hover:bg-slate-700 shadow-lg disabled:opacity-50 transition-all"
        >
          {saving ? 'Saving...' : saved ? '✓ Saved!' : 'Save Settings'}
        </button>
      </div>

      <div className="space-y-6">

        {/* ── Fare Percentages ─────────────────────────────────────────────── */}
        <Section title="Fare Percentages & Amounts">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <NumberField label="Half Fare (%)"          name="half_per"          value={formData.half_per}          onChange={handleChange} />
            <NumberField label="Concession (%)"         name="con_per"           value={formData.con_per}           onChange={handleChange} />
            <NumberField label="PH Concession (%)"      name="phy_per"           value={formData.phy_per}           onChange={handleChange} />
            <NumberField label="Student Max Amount"      name="st_max_amt"        value={formData.st_max_amt}        onChange={handleChange} />
            <NumberField label="Student Min Concession"  name="st_min_con"        value={formData.st_min_con}        onChange={handleChange} />
            <NumberField label="Rounding Amount"         name="round_amt"         value={formData.round_amt}         onChange={handleChange} />
            <NumberField label="Luggage Unit Rate"       name="luggage_unit_rate" value={formData.luggage_unit_rate} onChange={handleChange} />
            <NumberField label="ST Roundoff Amount"      name="st_roundoff_amt"   value={formData.st_roundoff_amt}   onChange={handleChange} />
          </div>
        </Section>

        {/* ── Ticket Display ───────────────────────────────────────────────── */}
        <Section title="Ticket Display (Headers & Footers)">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <TextField label="Main Display Line 1" name="main_display"  value={formData.main_display}  onChange={handleChange} />
            <TextField label="Main Display Line 2" name="main_display2" value={formData.main_display2} onChange={handleChange} />
            <TextField label="Header Line 1"       name="header1"       value={formData.header1}       onChange={handleChange} />
            <TextField label="Header Line 2"       name="header2"       value={formData.header2}       onChange={handleChange} />
            <TextField label="Header Line 3"       name="header3"       value={formData.header3}       onChange={handleChange} />
            <TextField label="Footer Line 1"       name="footer1"       value={formData.footer1}       onChange={handleChange} />
            <TextField label="Footer Line 2"       name="footer2"       value={formData.footer2}       onChange={handleChange} />
          </div>
        </Section>

        {/* ── Device & Passwords ───────────────────────────────────────────── */}
        <Section title="Device & Access">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <TextField label="Device ID (PalmtecID)" name="palmtec_id"  value={formData.palmtec_id}  onChange={handleChange} />
            <TextField label="User Password"          name="user_pwd"    value={formData.user_pwd}    onChange={handleChange} />
            <TextField label="Master Password"        name="master_pwd"  value={formData.master_pwd}  onChange={handleChange} />
            <TextField label="Currency"               name="currency"    value={formData.currency}    onChange={handleChange} placeholder="e.g. INR" />
            <NumberField label="Language Option"      name="language_option"    value={formData.language_option}    onChange={handleChange} />
            <NumberField label="Report Flag"          name="report_flag"        value={formData.report_flag}        onChange={handleChange} />
            <NumberField label="Report Font"          name="report_font"        value={formData.report_font}        onChange={handleChange} />
            <NumberField label="Default Stage"        name="default_stage"      value={formData.default_stage}      onChange={handleChange} />
            <NumberField label="Stage Updation Msg"   name="stage_updation_msg" value={formData.stage_updation_msg} onChange={handleChange} />
          </div>
        </Section>

        {/* ── Communication / FTP ──────────────────────────────────────────── */}
        <Section title="Communication & FTP">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <TextField label="Phone 2"        name="ph_no2"       value={formData.ph_no2}       onChange={handleChange} />
            <TextField label="Phone 3"        name="ph_no3"       value={formData.ph_no3}       onChange={handleChange} />
            <TextField label="Access Point"   name="access_point" value={formData.access_point} onChange={handleChange} />
            <TextField label="Destination"    name="dest_adds"    value={formData.dest_adds}    onChange={handleChange} />
            <TextField label="FTP Username"   name="username"     value={formData.username}     onChange={handleChange} />
            <TextField label="FTP Password"   name="password"     value={formData.password}     onChange={handleChange} />
            <TextField label="Upload Path"    name="uploadpath"   value={formData.uploadpath}   onChange={handleChange} />
            <TextField label="Download Path"  name="downloadpath" value={formData.downloadpath} onChange={handleChange} />
            <TextField label="HTTP URL"       name="http_url"     value={formData.http_url}     onChange={handleChange} />
          </div>
        </Section>

        {/* ── Feature Toggles ──────────────────────────────────────────────── */}
        <Section title="Feature Toggles">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
            <Toggle label="Round Off"              name="roundoff"              value={formData.roundoff}              onChange={handleChange} />
            <Toggle label="Round Up"               name="round_up"              value={formData.round_up}              onChange={handleChange} />
            <Toggle label="Remove Ticket Flag"     name="remove_ticket_flag"    value={formData.remove_ticket_flag}    onChange={handleChange} />
            <Toggle label="Stage Font Flag"        name="stage_font_flag"       value={formData.stage_font_flag}       onChange={handleChange} />
            <Toggle label="Next Fare Flag"         name="next_fare_flag"        value={formData.next_fare_flag}        onChange={handleChange} />
            <Toggle label="Odometer Entry"         name="odometer_entry"        value={formData.odometer_entry}        onChange={handleChange} />
            <Toggle label="Ticket No Big Font"     name="ticket_no_big_font"    value={formData.ticket_no_big_font}    onChange={handleChange} />
            <Toggle label="Crew Check"             name="crew_check"            value={formData.crew_check}            onChange={handleChange} />
            <Toggle label="GPRS Enable"            name="gprs_enable"           value={formData.gprs_enable}           onChange={handleChange} />
            <Toggle label="Trip Send"              name="tripsend_enable"       value={formData.tripsend_enable}       onChange={handleChange} />
            <Toggle label="Schedule Send"          name="schedulesend_enable"   value={formData.schedulesend_enable}   onChange={handleChange} />
            <Toggle label="Send Pending"           name="sendpend"              value={formData.sendpend}              onChange={handleChange} />
            <Toggle label="Inspector Report"       name="inspect_rpt"           value={formData.inspect_rpt}           onChange={handleChange} />
            <Toggle label="ST Roundoff"            name="st_roundoff_enable"    value={formData.st_roundoff_enable}    onChange={handleChange} />
            <Toggle label="ST Fare Edit"           name="st_fare_edit"          value={formData.st_fare_edit}          onChange={handleChange} />
            <Toggle label="Multiple Pass"          name="multiple_pass"         value={formData.multiple_pass}         onChange={handleChange} />
            <Toggle label="Simple Report"          name="simple_report"         value={formData.simple_report}         onChange={handleChange} />
            <Toggle label="Inspector SMS"          name="inspector_sms"         value={formData.inspector_sms}         onChange={handleChange} />
            <Toggle label="Auto Shutdown"          name="auto_shut_down"        value={formData.auto_shut_down}        onChange={handleChange} />
            <Toggle label="User Password Enable"   name="userpswd_enable"       value={formData.userpswd_enable}       onChange={handleChange} />
          </div>
        </Section>

        {/* Save button at bottom too for long page convenience */}
        <div className="flex justify-end">
          <button
            type="button" onClick={handleSave} disabled={saving}
            className="px-6 py-2.5 text-sm font-medium text-white bg-slate-800 rounded-xl hover:bg-slate-700 shadow-lg disabled:opacity-50 transition-all"
          >
            {saving ? 'Saving...' : saved ? '✓ Saved!' : 'Save Settings'}
          </button>
        </div>

      </div>
    </div>
  );
}
