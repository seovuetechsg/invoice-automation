'use client';

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';

export default function Home() {
  const [mounted, setMounted] = useState(false);
  
  // Tab states
  const [activeTab, setActiveTab] = useState('parser');
  
  // Rebate Module States
  const [chatPricesInput, setChatPricesInput] = useState('');
  const [rebateExceptionList, setRebateExceptionList] = useState([]);
  
  // Quarterly Reconciliation States
  const [reconFileStatus, setReconFileStatus] = useState('Upload payment list document/image to reconcile rebates.');
  const [reconFileColor, setReconFileColor] = useState('var(--text-muted)');
  const [reconInvoices, setReconInvoices] = useState([]);
  const [reconRecords, setReconRecords] = useState([]);
  const [reconLoading, setReconLoading] = useState(false);
  const [reconSearch, setReconSearch] = useState('');
  
  // Rebate Exception Manager States
  const [exceptionUploadStatus, setExceptionUploadStatus] = useState('Select an Excel (.xlsx) or CSV file to update exception model list.');
  const [exceptionUploadColor, setExceptionUploadColor] = useState('var(--text-muted)');
  const [exceptionSearch, setExceptionSearch] = useState('');
  
  // Database Profiles & Active Profile
  const [profiles, setProfiles] = useState([]);
  const [activeProfileName, setActiveProfileName] = useState('');
  const [activeProfile, setActiveProfile] = useState(null);
  
  // Reference schema fields discovered at runtime from ERPNext DocType
  const [poRefField, setPoRefField] = useState('supplier_refernce');
  const [piRefField, setPiRefField] = useState('bill_no');
  
  // Connection status
  const [connectionStatus, setConnectionStatus] = useState('disconnected'); // disconnected | connected | testing
  
  // Parser Settings (saved in localStorage)
  const [parserSettings, setParserSettings] = useState({
    type: 'regex',
    geminiKey: '',
    openaiKey: '',
    openaiModel: 'gpt-4o-mini'
  });
  
  // Form Settings state (for the currently selected profile editor)
  const [profileEditor, setProfileEditor] = useState({
    profile_name: '',
    url: '',
    auth_type: 'token',
    connection_type: 'proxy',
    api_key: '',
    api_secret: '',
    username: '',
    password: '',
    company: '',
    expense_account: '',
    cost_center: '',
    payment_account: '',
    creditors_account: '',
    sync_doctype: 'Purchase Order',
    warehouse: '',
    enable_sn_tracking: false
  });

  // Autocomplete cache lists
  const [suppliersList, setSuppliersList] = useState([]);
  const [itemsList, setItemsList] = useState([]);
  const [accountsList, setAccountsList] = useState([]);
  const [costCentersList, setCostCentersList] = useState([]);
  const [warehousesList, setWarehousesList] = useState([]);
  const [taxTemplatesList, setTaxTemplatesList] = useState([]);
  const [companiesList, setCompaniesList] = useState([]);
  const [loadingInvoices, setLoadingInvoices] = useState(false);
  const [autoSubmitFlow, setAutoSubmitFlow] = useState(false);
  
  // Autocomplete search states
  const [supplierQuery, setSupplierQuery] = useState('');
  const [supplierSuggestions, setSupplierSuggestions] = useState([]);
  const [showSupplierDropdown, setShowSupplierDropdown] = useState(false);
  const [selectedSupplierObject, setSelectedSupplierObject] = useState(null);

  // Execution logs
  const [logs, setLogs] = useState([]);
  
  // Progress states
  const [progress, setProgress] = useState({ show: false, percent: 0, title: '', details: '' });
  
  // Raw invoice text (direct PDF text or Tesseract OCR)
  const [rawText, setRawText] = useState('');
  const [showTextSection, setShowTextSection] = useState(false);
  
  // Verification Form Parsed Data
  const [invoiceForm, setInvoiceForm] = useState({
    invoice_number: '',
    date: '',
    supplier: '',
    grand_total: 0,
    discount_amount: 0,
    items: []
  });

  const [taxInclusive, setTaxInclusive] = useState(false);

  // Payment tab states
  const [paySupplierQuery, setPaySupplierQuery] = useState('');
  const [paySupplierSuggestions, setPaySupplierSuggestions] = useState([]);
  const [showPaySupplierDropdown, setShowPaySupplierDropdown] = useState(false);
  const [paySelectedSupplier, setPaySelectedSupplier] = useState(null);
  
  const [unpaidInvoices, setUnpaidInvoices] = useState([]);
  const [checkedInvoiceIds, setCheckedInvoiceIds] = useState({});
  const [paymentDate, setPaymentDate] = useState('');
  const [paymentMethod, setPaymentMethod] = useState('Bank');
  const [paymentAmount, setPaymentAmount] = useState(0);
  const [paymentListStatus, setPaymentListStatus] = useState('Select a supplier to load invoices.');
  const [paymentListColor, setPaymentListColor] = useState('var(--text-muted)');
  const [payInvoiceSearch, setPayInvoiceSearch] = useState('');
  const [paymentListCodes, setPaymentListCodes] = useState([]);
  const [matchedInvoiceSequence, setMatchedInvoiceSequence] = useState([]);
  const [paymentRefNo, setPaymentRefNo] = useState('');
  const [paymentRefDate, setPaymentRefDate] = useState('');
  const [paymentType, setPaymentType] = useState('Pay');

  // Dashboard logs history and metrics
  const [auditLogs, setAuditLogs] = useState([]);
  const [metrics, setMetrics] = useState({ total_count: 0, total_synced_amount: 0, total_paid_amount: 0 });

  // File inputs
  const fileInputRef = useRef(null);
  const payListInputRef = useRef(null);
  const logsEndRef = useRef(null);

  // Initial client setup
  useEffect(() => {
    setMounted(true);
    const today = new Date().toISOString().split('T')[0];
    setPaymentDate(today);
    setPaymentRefDate(today);
    
    // Load local settings
    const localParser = localStorage.getItem('parserSettings');
    if (localParser) {
      try {
        setParserSettings(JSON.parse(localParser));
      } catch (e) {
        console.error(e);
      }
    }

    // Load discovered reference fields
    const savedPoRef = localStorage.getItem('poRefField');
    if (savedPoRef) setPoRefField(savedPoRef);
    const savedPiRef = localStorage.getItem('piRefField');
    if (savedPiRef) setPiRefField(savedPiRef);
    
    // Fetch profiles & rebate exception list from Neon Postgres
    fetchProfilesAndHistory();
    fetchRebateExceptionList();
  }, []);

  async function fetchRebateExceptionList() {
    try {
      const res = await fetch('/api/rebates/exceptions');
      const data = await res.json();
      if (data.success && Array.isArray(data.data)) {
        setRebateExceptionList(data.data);
        logToConsole(`Loaded ${data.data.length} rebate exception model code(s) from Neon DB.`, 'info');
      }
    } catch (err) {
      console.error("Failed to fetch rebate exceptions:", err);
    }
  }

  // Sync scroll to bottom on console logs
  useEffect(() => {
    if (logsEndRef.current) {
      logsEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs]);

  // Handle active profile switching
  useEffect(() => {
    if (activeProfileName && profiles.length > 0) {
      const found = profiles.find(p => p.profile_name === activeProfileName);
      if (found) {
        const isRealSwitch = !activeProfile || activeProfile.profile_name !== activeProfileName;
        
        setActiveProfile(found);
        setProfileEditor(found);
        localStorage.setItem('activeProfileName', activeProfileName);
        
        if (isRealSwitch) {
          // Reset cached autocomplete items for the new site
          setSuppliersList([]);
          setItemsList([]);
          setAccountsList([]);
          setCostCentersList([]);
          setWarehousesList([]);
          setTaxTemplatesList([]);
          setCompaniesList([]);
          setSelectedSupplierObject(null);
          setSupplierQuery('');

          // Automatically fetch all ERPNext accounts and settings to populate dropdowns
          testProfileConnection(found);

          setPaySelectedSupplier(null);
          setPaySupplierQuery('');
          setUnpaidInvoices([]);
          setCheckedInvoiceIds({});
          setMatchedInvoiceSequence([]);
          setPaymentAmount(0);
          setPaymentListCodes([]);
          setPaymentRefNo('');
          const today = new Date().toISOString().split('T')[0];
          setPaymentRefDate(today);
          setPaymentType('Pay');
          setConnectionStatus('disconnected');
          
          logToConsole(`Switched active profile to: "${activeProfileName}"`, 'info');
        }
      }
    }
  }, [activeProfileName, profiles]);

  if (!mounted) return null;

  // Logging Helper
  function logToConsole(text, type = 'info') {
    setLogs(prev => [...prev, { text, type, timestamp: new Date().toLocaleTimeString() }]);
  }

  // Neon DB: Fetch profiles & action history
  async function fetchProfilesAndHistory() {
    try {
      const res = await fetch('/api/profiles');
      const data = await res.json();
      if (data.success) {
        // Merge with local credentials from LocalStorage so they are never saved in the shared DB
        const sanitizedProfiles = data.data.map(profile => {
          const localCredsStr = localStorage.getItem(`erp_creds_${profile.profile_name}`);
          if (localCredsStr) {
            try {
              const localCreds = JSON.parse(localCredsStr);
              return {
                ...profile,
                auth_type: localCreds.auth_type || profile.auth_type || 'token',
                api_key: localCreds.api_key || '',
                api_secret: localCreds.api_secret || '',
                username: localCreds.username || '',
                password: localCreds.password || ''
              };
            } catch (e) {
              console.error("Failed to parse local credentials:", e);
            }
          }
          return {
            ...profile,
            api_key: '',
            api_secret: '',
            username: '',
            password: ''
          };
        });

        setProfiles(sanitizedProfiles);
        
        // Restore active profile selection
        const savedProfile = localStorage.getItem('activeProfileName');
        if (savedProfile && sanitizedProfiles.some(p => p.profile_name === savedProfile)) {
          setActiveProfileName(savedProfile);
        } else if (sanitizedProfiles.length > 0) {
          setActiveProfileName(sanitizedProfiles[0].profile_name);
        }
      }
      
      // Fetch audit logs
      const historyRes = await fetch('/api/logs');
      const historyData = await historyRes.json();
      if (historyData.success) {
        setAuditLogs(historyData.data);
        setMetrics(historyData.metrics);
      }
    } catch (err) {
      logToConsole(`Error connecting to Neon database: ${err.message}`, 'err');
    }
  }

  // Neon DB: Log Action
  async function logAction(actionType, supplier, amount, docName, details = {}) {
    try {
      const res = await fetch('/api/logs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action_type: actionType,
          profile_name: activeProfileName,
          supplier,
          amount,
          doc_name: docName,
          details
        })
      });
      const data = await res.json();
      if (data.success) {
        // Refresh history
        fetchProfilesAndHistory();
      }
    } catch (err) {
      console.error("Failed to write audit log:", err);
    }
  }

  // ERPNext Server Proxy/Direct fetcher
  async function erpRequest(url, method = 'GET', data = null, overrideProfile = null) {
    const profile = overrideProfile || activeProfile;
    const profileName = overrideProfile ? overrideProfile.profile_name : activeProfileName;

    if (!profileName || !profile) {
      throw new Error("No active ERPNext profile selected.");
    }

    if (profile.connection_type === 'direct') {
      const targetUrl = `${profile.url.replace(/\/$/, '')}${url}`;
      const headers = {};
      
      if (profile.auth_type === 'token') {
        headers['Authorization'] = `token ${profile.api_key}:${profile.api_secret}`;
      } else {
        const token = btoa(`${profile.username}:${profile.password}`);
        headers['Authorization'] = `Basic ${token}`;
      }
      
      if (method && method !== 'GET') {
        headers['Content-Type'] = 'application/json';
      }

      console.log(`Direct Request: ${method} ${targetUrl}`);
      const res = await fetch(targetUrl, {
        method: method,
        headers: headers,
        credentials: 'include'
      });
      
      const resData = await res.json();
      if (!res.ok) {
        throw new Error(resData.message || `HTTP ${res.status}`);
      }
      return resData;
    }

    const res = await fetch('/api/erpnext', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        url,
        method,
        profile_name: profileName,
        profile: profile,
        data
      })
    });
    const resData = await res.json();
    if (!res.ok) {
      const errMsg = resData._server_messages 
        ? JSON.parse(resData._server_messages).map(m => JSON.parse(m).message).join(", ")
        : (resData.exception || resData.message || `HTTP ${res.status}`);
      throw new Error(errMsg);
    }
    return resData;
  }

  // Test ERPNext connection
  async function testProfileConnection(profile) {
    setConnectionStatus('testing');
    try {
      logToConsole(`Testing connection to ${profile.url}...`, 'info');
      const filterStr = JSON.stringify([["company_name", "=", profile.company || ""]]);
      
      let data;
      if (profile.connection_type === 'direct') {
        const targetUrl = `${profile.url.replace(/\/$/, '')}/api/resource/Company?filters=${encodeURIComponent(filterStr)}`;
        const headers = {};
        if (profile.auth_type === 'token') {
          headers['Authorization'] = `token ${profile.api_key}:${profile.api_secret}`;
        } else {
          headers['Authorization'] = `Basic ${btoa(profile.username + ':' + profile.password)}`;
        }
        
        console.log(`Direct Test Connection to: ${targetUrl}`);
        const res = await fetch(targetUrl, { method: 'GET', headers: headers, credentials: 'include' });
        data = await res.json();
        if (!res.ok) throw new Error(data.message || `HTTP ${res.status}`);
      } else {
        const res = await fetch('/api/erpnext', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            url: `/api/resource/Company?filters=${encodeURIComponent(filterStr)}`,
            method: 'GET',
            profile_name: profile.profile_name,
            profile: profile
          })
        });
        data = await res.json();
        if (!res.ok) throw new Error(data.error || data.exception || data.message || "Login failed");
      }
      
      if (data.data) {
        setConnectionStatus('connected');
        logToConsole(`✓ Connected successfully to "${profile.profile_name}"! Site: ${profile.url}`, 'success');
        
        // Discover reference field names defined on this ERPNext instance schema
        await discoverReferenceFields(profile);
        
        prefetchAutocompletes(profile);
      } else {
        setConnectionStatus('disconnected');
        logToConsole(`✗ Connection failed to "${profile.profile_name}". Please verify credentials.`, 'err');
      }
    } catch (err) {
      setConnectionStatus('disconnected');
      logToConsole(`✗ Connection error: ${err.message}`, 'err');
    }
  }

  // Discover DocType field schemas at runtime
  async function discoverReferenceFields(profile) {
    try {
      logToConsole("Discovering ERPNext reference fields schema...", "info");
      
      const getFields = async (doctype) => {
        try {
          const res = await fetch('/api/erpnext', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              url: `/api/method/frappe.desk.form.meta.get_meta?doctype=${encodeURIComponent(doctype)}`,
              method: 'GET',
              profile_name: profile.profile_name,
              profile: profile
            })
          });
          const data = await res.json();
          if (data && data.message && data.message.fields) {
            return data.message.fields.map(f => f.fieldname);
          }
        } catch (e) {
          try {
            const res = await fetch('/api/erpnext', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                url: `/api/resource/DocType/${encodeURIComponent(doctype)}`,
                method: 'GET',
                profile_name: profile.profile_name,
                profile: profile
              })
            });
            const data = await res.json();
            if (data && data.data && data.data.fields) {
              return data.data.fields.map(f => f.fieldname);
            }
          } catch (e2) {
            console.error(`DocType field query failed for ${doctype}:`, e2);
          }
        }
        return [];
      };

      // Discover for Purchase Order
      const poFields = await getFields('Purchase Order');
      if (poFields.length > 0) {
        let bestPOField = 'supplier_refernce';
        if (poFields.includes('supplier_refernce')) {
          bestPOField = 'supplier_refernce';
        } else if (poFields.includes('supplier_reference')) {
          bestPOField = 'supplier_reference';
        } else if (poFields.includes('bill_no')) {
          bestPOField = 'bill_no';
        } else {
          // Look for any custom field matching reference criteria
          const customRef = poFields.find(f => f.includes('invoice') || f.includes('ref') || f.includes('bill'));
          if (customRef) bestPOField = customRef;
        }
        setPoRefField(bestPOField);
        localStorage.setItem('poRefField', bestPOField);
        logToConsole(`✓ Selected Purchase Order reference field: "${bestPOField}"`, 'success');
      }

      // Discover for Purchase Invoice
      const piFields = await getFields('Purchase Invoice');
      if (piFields.length > 0) {
        let bestPIField = 'bill_no';
        if (piFields.includes('bill_no')) {
          bestPIField = 'bill_no';
        } else if (piFields.includes('supplier_invoice_no')) {
          bestPIField = 'supplier_invoice_no';
        } else if (piFields.includes('supplier_refernce')) {
          bestPIField = 'supplier_refernce';
        }
        setPiRefField(bestPIField);
        localStorage.setItem('piRefField', bestPIField);
        logToConsole(`✓ Selected Purchase Invoice reference field: "${bestPIField}"`, 'success');
      }
    } catch (err) {
      logToConsole(`Reference discovery warning: ${err.message}`, 'warn');
    }
  }

  // Pre-fetch Auto-completions
  async function prefetchAutocompletes(overrideProfile = null) {
    const profile = overrideProfile || activeProfile;
    if (!profile) return;
    
    try {
      // 1. Fetch Suppliers
      const fieldsSuppliers = JSON.stringify(["name", "supplier_name"]);
      const resSup = await erpRequest(`/api/resource/Supplier?fields=${encodeURIComponent(fieldsSuppliers)}&limit_page_length=1000`, 'GET', null, profile);
      if (resSup.data) {
        setSuppliersList(resSup.data);
        logToConsole(`Loaded ${resSup.data.length} suppliers for autocomplete cache.`, 'info');
      }

      // 2. Fetch Items in sequential batches of 2000 to prevent ERPNext server timeouts
      try {
        let allItems = [];
        let limit = 2000;
        let offset = 0;
        let hasMore = true;
        const maxItemsToLoad = 16000; // Safe ceiling to fetch up to 16,000 items
        
        logToConsole("Starting batch-loading of ERPNext items...", "info");
        
        while (hasMore && allItems.length < maxItemsToLoad) {
          const fieldsItems = JSON.stringify(["name", "item_name", "item_code", "last_purchase_rate"]);
          const resIt = await erpRequest(`/api/resource/Item?fields=${encodeURIComponent(fieldsItems)}&limit_page_length=${limit}&limit_start=${offset}`, 'GET', null, profile);
          
          if (resIt.data && resIt.data.length > 0) {
            allItems = [...allItems, ...resIt.data];
            logToConsole(`Loaded batch ${offset / limit + 1} (${resIt.data.length} items)...`, "info");
            offset += limit;
            if (resIt.data.length < limit) {
              hasMore = false; // reached end of database
            }
          } else {
            hasMore = false;
          }
        }
        
        if (allItems.length > 0) {
          setItemsList(allItems);
          logToConsole(`✓ Successfully loaded ${allItems.length} items with purchase rates into cache.`, 'success');
        } else {
          throw new Error("No items returned from ERPNext.");
        }
      } catch (itemErr) {
        logToConsole(`Failed to fetch items with purchase rates, falling back to basic fields. Error: ${itemErr.message}`, 'warn');
        try {
          let allItems = [];
          let limit = 2000;
          let offset = 0;
          let hasMore = true;
          const maxItemsToLoad = 16000;
          
          while (hasMore && allItems.length < maxItemsToLoad) {
            const fieldsItemsBasic = JSON.stringify(["name", "item_name", "item_code"]);
            const resIt = await erpRequest(`/api/resource/Item?fields=${encodeURIComponent(fieldsItemsBasic)}&limit_page_length=${limit}&limit_start=${offset}`, 'GET', null, profile);
            
            if (resIt.data && resIt.data.length > 0) {
              allItems = [...allItems, ...resIt.data];
              offset += limit;
              if (resIt.data.length < limit) {
                hasMore = false;
              }
            } else {
              hasMore = false;
            }
          }
          if (allItems.length > 0) {
            setItemsList(allItems);
            logToConsole(`✓ Loaded ${allItems.length} items (basic fields) into cache.`, 'success');
          }
        } catch (fbErr) {
          logToConsole(`Fallback basic item pre-fetch failed: ${fbErr.message}`, 'err');
        }
      }

      // 3. Fetch Accounts
      try {
        const filterStrAcc = JSON.stringify([["is_group", "=", 0], ["docstatus", "<", 2]]);
        const resAcc = await erpRequest(`/api/resource/Account?filters=${encodeURIComponent(filterStrAcc)}&fields=${encodeURIComponent(JSON.stringify(["name"]))}&limit_page_length=3000`, 'GET', null, profile);
        if (resAcc.data) {
          setAccountsList(resAcc.data);
          logToConsole(`Loaded ${resAcc.data.length} accounts from ERPNext for settings autocompletion.`, 'info');
        }
      } catch (accErr) {
        logToConsole(`Accounts autocompletion fetch failed: ${accErr.message}`, 'warn');
      }

      // 4. Fetch Cost Centers
      try {
        const filterStrCc = JSON.stringify([["is_group", "=", 0], ["docstatus", "<", 2]]);
        const resCc = await erpRequest(`/api/resource/Cost Center?filters=${encodeURIComponent(filterStrCc)}&fields=${encodeURIComponent(JSON.stringify(["name"]))}&limit_page_length=1000`, 'GET', null, profile);
        if (resCc.data) {
          setCostCentersList(resCc.data);
          logToConsole(`Loaded ${resCc.data.length} cost centers from ERPNext for settings autocompletion.`, 'info');
        }
      } catch (ccErr) {
        logToConsole(`Cost Centers autocompletion fetch failed: ${ccErr.message}`, 'warn');
      }

      // 5. Fetch Warehouses
      try {
        const filterStrWh = JSON.stringify([["is_group", "=", 0], ["docstatus", "<", 2]]);
        const resWh = await erpRequest(`/api/resource/Warehouse?filters=${encodeURIComponent(filterStrWh)}&fields=${encodeURIComponent(JSON.stringify(["name"]))}&limit_page_length=1000`, 'GET', null, profile);
        if (resWh.data) {
          setWarehousesList(resWh.data);
          logToConsole(`Loaded ${resWh.data.length} warehouses from ERPNext for settings autocompletion.`, 'info');
        }
      } catch (whErr) {
        logToConsole(`Warehouses autocompletion fetch failed: ${whErr.message}`, 'warn');
      }

      // 6. Fetch Purchase Taxes and Charges Templates
      try {
        const resTax = await erpRequest(`/api/resource/Purchase Taxes and Charges Template?fields=${encodeURIComponent(JSON.stringify(["name"]))}&limit_page_length=1000`, 'GET', null, profile);
        if (resTax.data) {
          setTaxTemplatesList(resTax.data);
          logToConsole(`Loaded ${resTax.data.length} tax templates from ERPNext for settings autocompletion.`, 'info');
        }
      } catch (taxErr) {
        logToConsole(`Tax templates autocompletion fetch failed: ${taxErr.message}`, 'warn');
      }

      // 7. Fetch Companies
      try {
        const resComp = await erpRequest(`/api/resource/Company?fields=${encodeURIComponent(JSON.stringify(["name"]))}&limit_page_length=100`, 'GET', null, profile);
        if (resComp.data) {
          setCompaniesList(resComp.data);
          logToConsole(`Loaded ${resComp.data.length} companies from ERPNext for settings autocompletion.`, 'info');
        }
      } catch (compErr) {
        logToConsole(`Companies autocompletion fetch failed: ${compErr.message}`, 'warn');
      }
    } catch (e) {
      logToConsole(`Autocomplete pre-fetch failed: ${e.message}`, 'warn');
    }
  }

  // Profile Editor Actions
  const handleEditorChange = (e) => {
    const { id, value } = e.target;
    const field = id.replace('erp-', '');
    setProfileEditor(prev => {
      const next = { ...prev, [field]: value };
      
      // Clear alternative auth details when type switches
      if (field === 'auth_type') {
        if (value === 'token') {
          next.username = '';
          next.password = '';
        } else if (value === 'password') {
          next.api_key = '';
          next.api_secret = '';
        }
      }
      return next;
    });
  };

  async function saveProfileSettings() {
    if (!profileEditor.profile_name || !profileEditor.url) {
      alert("Profile name and Site URL are required.");
      return;
    }
    
    try {
      logToConsole(`Saving settings for profile "${profileEditor.profile_name}" in database...`, 'info');
      
      // 1. Save sensitive credentials locally in the browser's localStorage
      const localCreds = {
        auth_type: profileEditor.auth_type || 'token',
        api_key: profileEditor.api_key || '',
        api_secret: profileEditor.api_secret || '',
        username: profileEditor.username || '',
        password: profileEditor.password || ''
      };
      localStorage.setItem(`erp_creds_${profileEditor.profile_name}`, JSON.stringify(localCreds));
      
      // 2. Prepare database payload with cleared credential fields
      const payload = { 
        ...profileEditor, 
        sync_doctype: 'Purchase Order',
        api_key: '',
        api_secret: '',
        username: '',
        password: ''
      };
      
      const res = await fetch('/api/profiles', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data.success) {
        logToConsole(`✓ Settings saved successfully for "${profileEditor.profile_name}"!`, 'success');
        await fetchProfilesAndHistory();
        setActiveProfileName(profileEditor.profile_name);
      } else {
        throw new Error(data.error);
      }
    } catch (err) {
      logToConsole(`Error saving profile: ${err.message}`, 'err');
      alert(`Save failed: ${err.message}`);
    }
  }

  async function deleteActiveProfile() {
    if (!activeProfileName) return;
    if (!confirm(`Are you sure you want to delete profile "${activeProfileName}"?`)) return;

    try {
      const res = await fetch(`/api/profiles?profile_name=${encodeURIComponent(activeProfileName)}`, {
        method: 'DELETE'
      });
      const data = await res.json();
      if (data.success) {
        logToConsole(`✓ Profile "${activeProfileName}" deleted from database.`, 'success');
        
        // Remove locally cached credentials
        localStorage.removeItem(`erp_creds_${activeProfileName}`);
        
        setActiveProfileName('');
        setActiveProfile(null);
        setProfileEditor({
          profile_name: '', url: '', auth_type: 'token', api_key: '', api_secret: '',
          username: '', password: '', company: '', expense_account: '', cost_center: '',
          payment_account: '', creditors_account: '', sync_doctype: 'Purchase Order',
          warehouse: ''
        });
        await fetchProfilesAndHistory();
      } else {
        throw new Error(data.error);
      }
    } catch (err) {
      logToConsole(`Delete failed: ${err.message}`, 'err');
      alert(`Delete failed: ${err.message}`);
    }
  }

  function handleCreateNewProfile() {
    const pName = prompt("Enter a unique name for the new ERPNext Profile:");
    if (!pName) return;
    const cleanName = pName.trim();
    if (profiles.some(p => p.profile_name.toLowerCase() === cleanName.toLowerCase())) {
      alert("A profile with this name already exists.");
      return;
    }
    
    const newTemplate = {
      profile_name: cleanName,
      url: 'https://',
      auth_type: 'token',
      connection_type: 'proxy',
      api_key: '',
      api_secret: '',
      username: '',
      password: '',
      company: '',
      expense_account: 'Cost of Goods Sold - CEF',
      cost_center: 'Main - CEF',
      payment_account: 'Bank Account - CEF',
      creditors_account: 'Creditors - CEF',
      sync_doctype: 'Purchase Order',
      warehouse: 'Stores - CEF'
    };

    setProfiles(prev => [...prev, newTemplate]);
    setActiveProfileName(cleanName);
    setProfileEditor(newTemplate);
  }

  // Parse Settings Engine actions
  function saveParserSettings(newSettings) {
    setParserSettings(newSettings);
    localStorage.setItem('parserSettings', JSON.stringify(newSettings));
    logToConsole("AI Parser engine configuration updated.", 'success');
    alert("AI Parser settings saved successfully!");
  }

  // PDF.js / OCR File Loading Pipeline
  const triggerFileSelect = () => fileInputRef.current.click();

  const handleFileDropChange = async (e) => {
    const files = e.target.files || e.dataTransfer?.files;
    if (!files || files.length === 0) return;
    const file = files[0];
    
    if (!file.type.startsWith('image/') && file.type !== 'application/pdf') {
      logToConsole("Unsupported file type. Please upload an image or PDF.", 'warn');
      return;
    }
    
    logToConsole(`Loading file: ${file.name} (${Math.round(file.size / 1024)} KB)...`, 'info');
    setProgress({ show: true, percent: 5, title: 'Reading PDF...', details: 'Analyzing page structure...' });
    
    if (file.type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        processImageSource(evt.target.result);
      };
      reader.readAsDataURL(file);
    } else if (file.type === 'application/pdf') {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        try {
          const arrayBuffer = evt.target.result;
          
          // Dynamically import PDF.js client-side
          const pdfjsLib = await import('pdfjs-dist');
          pdfjsLib.GlobalWorkerOptions.workerSrc = `//unpkg.com/pdfjs-dist@6.1.200/build/pdf.worker.min.mjs`;

          const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
          const pdf = await loadingTask.promise;
          logToConsole(`✓ PDF loaded. Total pages: ${pdf.numPages}. Checking for selectable text...`, 'info');
          
          const page = await pdf.getPage(1);
          const textContent = await page.getTextContent();
          const items = textContent.items || [];
          
          let extractedText = '';
          if (items.length > 0) {
            // Sort items top-to-bottom, left-to-right
            items.sort((a, b) => {
              const yDiff = b.transform[5] - a.transform[5];
              if (Math.abs(yDiff) > 5) return yDiff;
              return a.transform[4] - b.transform[4];
            });
            
            let lastY = null;
            for (const item of items) {
              if (lastY !== null && Math.abs(item.transform[5] - lastY) > 5) {
                extractedText += '\n';
              } else if (lastY !== null) {
                extractedText += ' ';
              }
              extractedText += item.str;
              lastY = item.transform[5];
            }
            extractedText = extractedText.trim();
          }
          
          if (extractedText.length > 10) {
            logToConsole("✓ Digital PDF detected. Selectable text extracted directly (100% accurate).", 'success');
            setProgress({ show: true, percent: 30, title: 'Parsing Text...', details: 'Extracting invoice fields...' });
            
            setRawText(extractedText);
            setShowTextSection(true);
            parseInvoiceText(extractedText);
            return;
          }
          
          // Fallback: Scanned PDF
          logToConsole("⚠ Scanned PDF detected (no selectable text). Falling back to image rendering...", 'warn');
          setProgress({ show: true, percent: 15, title: 'Rendering PDF...', details: 'Converting page 1 to image...' });
          
          const scale = 2.0;
          const viewport = page.getViewport({ scale });
          const canvas = document.createElement('canvas');
          const context = canvas.getContext('2d');
          canvas.height = viewport.height;
          canvas.width = viewport.width;
          
          await page.render({ canvasContext: context, viewport }).promise;
          const dataUrl = canvas.toDataURL('image/png');
          logToConsole("✓ Page 1 rendered successfully. Initializing OCR/Vision...", 'info');
          processImageSource(dataUrl);
          
        } catch (error) {
          setProgress(prev => ({ ...prev, show: false }));
          logToConsole(`PDF Processing failed: ${error.message || error}`, 'err');
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // Process Images (Vision or local Tesseract)
  async function processImageSource(imageSrc) {
    const engineType = parserSettings.type;
    
    if (engineType === 'regex') {
      try {
        logToConsole("Initializing Tesseract.js worker locally...", 'info');
        setProgress({ show: true, percent: 25, title: 'Starting Tesseract...', details: 'Loading offline language model...' });
        
        // Dynamically import tesseract client-side
        const { createWorker } = await import('tesseract.js');
        const worker = await createWorker('eng');
        setProgress({ show: true, percent: 45, title: 'Scanning Image...', details: 'Recognizing letters and coordinates...' });
        
        const result = await worker.recognize(imageSrc);
        const text = result.data.text;
        logToConsole("✓ Local OCR scan finished successfully.", 'success');
        
        await worker.terminate();
        
        setRawText(text);
        setShowTextSection(true);
        parseInvoiceText(text);
        
      } catch (error) {
        setProgress(prev => ({ ...prev, show: false }));
        logToConsole(`Local OCR failure: ${error.message || error}`, 'err');
        console.error(error);
      }
    } else {
      // Cloud AI Vision Models
      try {
        setProgress({ show: true, percent: 40, title: 'Calling Cloud AI...', details: `Uploading image to ${engineType === 'gemini' ? 'Google Gemini' : 'OpenAI'}...` });
        logToConsole(`Uploading invoice image to ${engineType === 'gemini' ? 'Google Gemini' : 'OpenAI'} Cloud Vision...`, 'info');
        
        const base64Parts = imageSrc.split(',');
        const mimeType = base64Parts[0].match(/:(.*?);/)[1];
        const base64Data = base64Parts[1];
        
        let parsedJson = null;
        
        if (engineType === 'gemini') {
          if (!parserSettings.geminiKey) {
            throw new Error("Google Gemini API Key is missing. Configure in settings.");
          }
          const cleanKey = parserSettings.geminiKey.trim();
          
          const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${cleanKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{
                parts: [
                  { text: `Extract structured data from this invoice. Identify the Supplier/Vendor name, Tax Invoice No (invoice_number). If no invoice/bill number is detected, extract the Sales Order (SO) number instead for the 'invoice_number' field. Also extract Invoice Date (date), Grand Total (grand_total), and all Line Items. For the description of each line item, extract ONLY the model code, model name, or item code (e.g. 'MX-200', 'ITEM-101'). Do NOT include features, specs, product descriptions, or long text. Keep it very short. If discount exists, extract rate as unit price minus discount. CRITICAL: Scan the text directly underneath each item's description for labels like 'Serial No:', 'S/N:', or lists of long numbers. Extract those exact serial numbers as a comma-separated string in 'serial_nos' (leave empty if none). Respond ONLY with a valid JSON matching this schema: {\"invoice_number\":\"...\",\"date\":\"YYYY-MM-DD\",\"supplier\":\"...\",\"grand_total\":0.0,\"items\":[{\"description\":\"...\",\"qty\":1,\"rate\":0.0,\"amount\":0.0,\"serial_nos\":\"SN1, SN2\"}]}` },
                  { inlineData: { mimeType, data: base64Data } }
                ]
              }],
              generationConfig: { responseMimeType: 'application/json' }
            })
          });
          
          if (!response.ok) {
            const errData = await response.json().catch(() => null);
            const errMsg = errData?.error?.message || `HTTP ${response.status}`;
            throw new Error(`Gemini API failed: ${errMsg}`);
          }
          const resData = await response.json();
          const textResponse = resData.candidates[0].content.parts[0].text.trim();
          parsedJson = JSON.parse(textResponse);
          
        } else if (engineType === 'openai') {
          if (!parserSettings.openaiKey) {
            throw new Error("OpenAI API Key is missing. Configure in settings.");
          }
          
          const response = await fetch("https://api.openai.com/v1/chat/completions", {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${parserSettings.openaiKey}`
            },
            body: JSON.stringify({
              model: parserSettings.openaiModel || 'gpt-4o-mini',
              response_format: { type: 'json_object' },
              messages: [
                {
                  role: 'system',
                  content: "You are an expert invoice visual parser. For the 'invoice_number' field, if no invoice/bill number is detected in the text, extract the Sales Order (SO) number instead as a fallback. For the description of each line item, extract ONLY the model code, model name, or item code (e.g. 'MX-200', 'ITEM-101'). Do NOT include features, specifications, or descriptions. Keep it very short. CRITICAL: Scan the text directly underneath each item's description for labels like 'Serial No:', 'S/N:', or lists of long numbers. Extract those exact serial numbers as a comma-separated string in 'serial_nos' (leave empty if none). Respond ONLY with a valid JSON object matching this schema:\n{\n  \"invoice_number\": \"...\",\n  \"date\": \"YYYY-MM-DD\",\n  \"supplier\": \"...\",\n  \"grand_total\": 0.0,\n  \"items\": [\n    { \"description\": \"...\", \"qty\": 1, \"rate\": 0.0, \"amount\": 0.0, \"serial_nos\": \"\" }\n  ]\n}"
                },
                {
                  role: 'user',
                  content: [
                    { type: 'text', text: "Extract structured data from this invoice. Description must strictly be the model name/code." },
                    { type: 'image_url', image_url: { url: imageSrc } }
                  ]
                }
              ]
            })
          });
          
          if (!response.ok) throw new Error(`OpenAI API failed: HTTP ${response.status}`);
          const resData = await response.json();
          parsedJson = JSON.parse(resData.choices[0].message.content.trim());
        }
        
        logToConsole(`✓ Cloud Vision extraction completed successfully via ${engineType}.`, 'success');
        setProgress(prev => ({ ...prev, show: false }));
        processRawParsedData(parsedJson);
        
      } catch (error) {
        setProgress(prev => ({ ...prev, show: false }));
        logToConsole(`Cloud Vision error: ${error.message || error}`, 'err');
      }
    }
  }

  // Parse direct text
  async function parseInvoiceText(text) {
    const type = parserSettings.type;
    
    if (type === 'gemini' || type === 'openai') {
      try {
        setProgress({ show: true, percent: 40, title: 'Calling Cloud AI...', details: `Parsing text via ${type === 'gemini' ? 'Google Gemini' : 'OpenAI'}...` });
        logToConsole(`Sending raw extracted text to ${type === 'gemini' ? 'Google Gemini' : 'OpenAI'} API...`, 'info');
        
        let parsedJson = null;
        
        if (type === 'gemini') {
          if (!parserSettings.geminiKey) throw new Error("Gemini API key is missing.");
          const cleanKey = parserSettings.geminiKey.trim();
          
          const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${cleanKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{
                parts: [{
                  text: `Extract structured data from this raw invoice text. Look closely for the Supplier/Vendor name, Tax Invoice No (invoice_number). If no invoice/bill number is detected, extract the Sales Order (SO) number instead for the 'invoice_number' field. Also find Invoice Date (date), Grand Total (grand_total), and all Line Items. For the description of each line item, extract ONLY the model code, model name, or item code (e.g. 'MX-200', 'ITEM-101'). Do NOT include features, product descriptions, specs, or long marketing texts. Keep it very short. If discount exists, extract rate as unit price minus discount. CRITICAL: Scan the text directly underneath each item's description for labels like 'Serial No:', 'S/N:', or lists of long numbers. Extract those exact serial numbers as a comma-separated string in 'serial_nos' (leave empty if none). Respond ONLY with a valid JSON matching this schema: {\"invoice_number\":\"...\",\"date\":\"YYYY-MM-DD\",\"supplier\":\"...\",\"grand_total\":0.0,\"items\":[{\"description\":\"...\",\"qty\":1,\"rate\":0.0,\"amount\":0.0,\"serial_nos\":\"\"}]}\n\nRaw Invoice Text:\n${text}`
                }]
              }],
              generationConfig: { responseMimeType: 'application/json' }
            })
          });
          
          if (!response.ok) {
            const errData = await response.json().catch(() => null);
            const errMsg = errData?.error?.message || `HTTP ${response.status}`;
            throw new Error(`Gemini status: ${errMsg}`);
          }
          const resData = await response.json();
          parsedJson = JSON.parse(resData.candidates[0].content.parts[0].text.trim());
          
        } else if (type === 'openai') {
          if (!parserSettings.openaiKey) throw new Error("OpenAI API key is missing.");
          const response = await fetch("https://api.openai.com/v1/chat/completions", {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${parserSettings.openaiKey}`
            },
            body: JSON.stringify({
              model: parserSettings.openaiModel || 'gpt-4o-mini',
              response_format: { type: 'json_object' },
              messages: [
                {
                  role: 'system',
                  content: "You are an expert invoice text parser. For the 'invoice_number' field, if no invoice/bill number is detected in the text, extract the Sales Order (SO) number instead as a fallback. For the description of each line item, extract ONLY the model code, model name, or item code (e.g. 'MX-200', 'ITEM-101'). Do NOT include product features, specifications, or long descriptions. Keep it very short. CRITICAL: Scan the text directly underneath each item's description for labels like 'Serial No:', 'S/N:', or lists of long numbers. Extract those exact serial numbers as a comma-separated string in 'serial_nos' (leave empty if none). Respond ONLY with a valid JSON object matching this schema:\n{\n  \"invoice_number\": \"...\",\n  \"date\": \"YYYY-MM-DD\",\n  \"supplier\": \"...\",\n  \"grand_total\": 0.0,\n  \"items\": [\n    { \"description\": \"...\", \"qty\": 1, \"rate\": 0.0, \"amount\": 0.0, \"serial_nos\": \"\" }\n  ]\n}"
                },
                {
                  role: 'user',
                  content: `Extract structured data from this raw invoice text. Focus on Supplier name, Tax Invoice No, Date, Grand Total, and Line Items. Ensure description is strictly the model name or code.\n\nRaw Invoice Text:\n${text}`
                }
              ]
            })
          });
          
          if (!response.ok) throw new Error(`OpenAI status ${response.status}`);
          const resData = await response.json();
          parsedJson = JSON.parse(resData.choices[0].message.content.trim());
        }
        
        logToConsole(`✓ Cloud text parsing completed successfully via ${type}.`, 'success');
        setProgress(prev => ({ ...prev, show: false }));
        processRawParsedData(parsedJson);
        
      } catch (error) {
        setProgress(prev => ({ ...prev, show: false }));
        logToConsole(`Cloud Text Parser Error: ${error.message || error}`, 'err');
      }
    } else {
      parseWithRegex(text);
    }
  }

  // Filter discounts and trigger verification form rendering
  function processRawParsedData(parsedJson) {
    let discountAmount = 0;
    const filteredItems = [];
    
    (parsedJson.items || []).forEach(i => {
      const desc = (i.description || i.item_name || "").toLowerCase();
      const qty = parseFloat(i.qty || i.quantity || 1);
      const rate = parseFloat(i.rate || i.price || 0);
      const amount = parseFloat(i.amount || (qty * rate) || 0);
      
      const isDiscountText = desc.includes("discount") || desc.includes("rebate") || desc.includes("allowance") || desc.includes("less:") || desc.includes("promo");
      const isNegative = rate < 0 || amount < 0;
      
      if (isDiscountText || isNegative) {
        discountAmount += Math.abs(amount || rate);
      } else {
        const matchResult = findMatchingItemCode(i.description || "Line Item");
        filteredItems.push({
          description: i.description || "Line Item",
          qty: qty,
          rate: rate,
          amount: amount,
          item_code: matchResult.item_code,
          match_confidence: matchResult.match_confidence,
          serial_nos: i.serial_nos || ""
        });
      }
    });

    const adaptedData = {
      invoice_number: parsedJson.invoice_number || "",
      date: formatDate(parsedJson.date || ""),
      supplier: parsedJson.supplier || "",
      grand_total: parseFloat(parsedJson.grand_total || 0),
      discount_amount: discountAmount,
      items: filteredItems
    };
    
    populateVerificationForm(adaptedData);
  }

  // Regex Extraction Layer
  function parseWithRegex(text) {
    logToConsole("Parsing invoice text using local rule-based algorithms...", 'info');
    
    const parsedData = {
      invoice_number: "",
      date: "",
      supplier: "",
      grand_total: 0.0,
      items: []
    };
    
    const lines = text.split("\n");
    
    // Invoice No
    const invNumRegexes = [
      /(?:tax\s*invoice\s*no|tax\s*invoice|invoice\s*no|invoice\s*number|inv\s*#|invoice\s*#|bill\s*no|reference\s*no)[\s:#]*([A-Z0-9-_]+)/i,
      /(?:tax\s*invoice\s*number|invoice\s*id)[\s:#]*([A-Z0-9-_]+)/i
    ];
    for (const line of lines) {
      let matched = false;
      for (const regex of invNumRegexes) {
        const match = line.match(regex);
        if (match && match[1]) {
          parsedData.invoice_number = match[1].trim();
          matched = true;
          break;
        }
      }
      if (matched) break;
    }

    // Fallback: Sales Order number if no invoice number was detected
    if (!parsedData.invoice_number) {
      const salesOrderRegexes = [
        /(?:sales\s*order\s*no|sales\s*order\s*#|sales\s*order\s*number|sales\s*order|order\s*no|order\s*number|so\s*no|so\s*#|s\.o\.\s*no|s\/o)[\s:#]*([A-Z0-9-_]+)/i
      ];
      for (const line of lines) {
        let matched = false;
        for (const regex of salesOrderRegexes) {
          const match = line.match(regex);
          if (match && match[1]) {
            parsedData.invoice_number = match[1].trim();
            matched = true;
            break;
          }
        }
        if (matched) break;
      }
    }
    
    // Date
    const dateRegex = /\b(\d{1,4}[-\/\.]\d{1,2}[-\/\.]\d{1,4})\b/;
    for (const line of lines) {
      const match = line.match(dateRegex);
      if (match && match[1]) {
        parsedData.date = formatDate(match[1]);
        break;
      }
    }
    
    // Grand Total
    const totalRegex = /(?:grand\s*total|total\s*payable|total|payable\s*amount|amount\s*due|net\s*payable)[\s:#]*\$?\s*([\d,]+\.\d{2})/i;
    for (const line of lines) {
      const match = line.match(totalRegex);
      if (match && match[1]) {
        parsedData.grand_total = parseFloat(match[1].replace(/,/g, ''));
        break;
      }
    }
    
    // Extract potential table line items
    const itemRowRegex = /^(.*?)\b(\d+)\s+[\$]?\s*([\d,]+\.\d{2})\s+[\$]?\s*([\d,]+\.\d{2})\s*$/;
    for (const line of lines) {
      const match = line.match(itemRowRegex);
      if (match) {
        const desc = match[1].trim();
        const qty = parseFloat(match[2]);
        const rate = parseFloat(match[3].replace(/,/g, ''));
        const amount = parseFloat(match[4].replace(/,/g, ''));
        
        const descLower = desc.toLowerCase();
        const isDiscount = descLower.includes("discount") || descLower.includes("rebate") || descLower.includes("allowance") || rate < 0 || amount < 0;
        
        if (isDiscount) {
          parsedData.discount_amount = (parsedData.discount_amount || 0) + Math.abs(amount || rate);
        } else if (desc.length > 2 && qty > 0 && rate > 0) {
          const matchResult = findMatchingItemCode(desc);
          parsedData.items.push({
            description: desc,
            qty,
            rate,
            amount,
            item_code: matchResult.item_code,
            match_confidence: matchResult.match_confidence
          });
        }
      }
    }
    
    logToConsole("✓ Rule-based extraction finished.", 'success');
    setProgress(prev => ({ ...prev, show: false }));
    populateVerificationForm(parsedData);
  }

  // Populate Verification form & Match supplier
  function populateVerificationForm(data) {
    const query = (data.supplier || "").toLowerCase().trim();
    let bestMatch = "";
    
    if (query && suppliersList.length > 0) {
      let maxScore = 0;
      suppliersList.forEach(s => {
        const sName = (s.supplier_name || s.name || "").toLowerCase();
        if (sName.includes(query) || query.includes(sName)) {
          const score = Math.max(sName.length, query.length);
          if (score > maxScore) {
            maxScore = score;
            bestMatch = s.name;
          }
        }
      });
    }

    const consolidatedItems = [];
    if (data.items) {
      const tempMap = {};
      data.items.forEach(line => {
        const itemCode = line.item_code || 'Default Item';
        const rate = parseFloat(line.rate || 0);
        const qty = parseFloat(line.qty || 0);
        const desc = line.description || '';
        const key = `${itemCode}_${rate.toFixed(4)}`;
        if (tempMap[key]) {
          tempMap[key].qty += qty;
          tempMap[key].amount = tempMap[key].qty * tempMap[key].rate;
          if (desc && !tempMap[key].description.includes(desc)) {
            tempMap[key].description += ` / ${desc}`;
          }
        } else {
          tempMap[key] = {
            ...line,
            item_code: itemCode,
            qty: qty,
            rate: rate,
            amount: qty * rate,
            description: desc
          };
        }
      });
      consolidatedItems.push(...Object.values(tempMap));
    }

    setInvoiceForm({
      invoice_number: data.invoice_number || '',
      date: data.date || '',
      supplier: bestMatch || data.supplier || '',
      grand_total: data.grand_total || 0,
      discount_amount: data.discount_amount || 0,
      items: consolidatedItems
    });
    setPaymentListCodes(data.invoice_number ? [data.invoice_number] : []);
    
    setSelectedSupplierObject(suppliersList.find(s => s.name === bestMatch) || null);
    setSupplierQuery(bestMatch ? (suppliersList.find(s => s.name === bestMatch).supplier_name || bestMatch) : data.supplier || '');
    setShowSupplierDropdown(false);
    
    logToConsole("Verification Form populated with extracted values.", 'info');
    setActiveTab('parser');
  }

  // Sync to ERPNext Dispatcher
  async function syncInvoiceToERPNext() {
    if (connectionStatus !== 'connected') {
      alert("ERPNext connection is offline. Configure profile first.");
      return;
    }
    
    if (!selectedSupplierObject) {
      alert("Supplier must be linked to an ERPNext supplier record.");
      return;
    }
    
    // Check for missing items in ERPNext and prompt for auto-creation
    const missingItems = [];
    for (const line of invoiceForm.items) {
      const isDiscount = parseFloat(line.rate || 0) < 0 || (line.description || '').toLowerCase().includes('discount');
      if (isDiscount) continue;
      
      const itemCode = (line.item_code || line.description || '').trim();
      if (!itemCode || itemCode === 'Default Item') continue;

      const erpItemMatch = itemsList.find(it => (it.item_code || it.name || '').toLowerCase() === itemCode.toLowerCase());
      if (!erpItemMatch && !missingItems.includes(itemCode)) {
        missingItems.push(itemCode);
      }
    }

    if (missingItems.length > 0) {
      const confirmCreate = confirm(`⚠️ The following item(s) are NOT found in ERPNext:\n\n${missingItems.map(m => `• ${m}`).join('\n')}\n\nWould you like the app to automatically create them as new Products before syncing?`);
      if (confirmCreate) {
        logToConsole(`Auto-creating ${missingItems.length} missing item(s) in ERPNext...`, 'info');
        let createdAny = false;
        
        for (const missingCode of missingItems) {
          try {
            const newItemPayload = {
              item_code: missingCode,
              item_name: missingCode,
              item_group: "Product", // User requested 'Product'
              stock_uom: "Nos",
              is_stock_item: 1
            };
            
            // Try to create the item
            const createRes = await erpRequest('/api/resource/Item', 'POST', newItemPayload);
            if (createRes && createRes.data) {
              logToConsole(`✓ Successfully created item: ${missingCode}`, 'success');
              // Append to local state cache immediately so it passes mapping later
              setItemsList(prev => [...prev, createRes.data]);
              createdAny = true;
            }
          } catch (createErr) {
            // Fallback for strict Item Group naming in standard ERPNext
            if (createErr.message.includes('Item Group')) {
               logToConsole(`Retrying ${missingCode} with item_group="Products"...`, 'warn');
               try {
                 const fallbackPayload = { item_code: missingCode, item_name: missingCode, item_group: "Products", stock_uom: "Nos", is_stock_item: 1 };
                 const fbRes = await erpRequest('/api/resource/Item', 'POST', fallbackPayload);
                 if (fbRes && fbRes.data) {
                   logToConsole(`✓ Successfully created item: ${missingCode}`, 'success');
                   setItemsList(prev => [...prev, fbRes.data]);
                   createdAny = true;
                 }
               } catch (fbErr) {
                 alert(`Failed to create item ${missingCode}. Check Item Group or mandatory fields in ERPNext.`);
                 return;
               }
            } else {
              alert(`Failed to create item ${missingCode}: ${createErr.message}`);
              return;
            }
          }
        }
        if (createdAny) {
           logToConsole(`✓ Auto-creation complete. Proceeding with sync...`, 'success');
        }
      } else {
        logToConsole(`Sync cancelled by user (pending missing items).`, 'warn');
        return;
      }
    }
    

    
    const docType = activeProfile.sync_doctype || 'Purchase Order';
    logToConsole(`Initiating sync: Creating Draft ${docType} in ERPNext...`, 'info');
    
    try {
      const consolidatedMap = {};
      for (const line of invoiceForm.items) {
        const itemCode = line.item_code || 'Default Item';
        const rate = parseFloat(line.rate || 0);
        const qty = parseFloat(line.qty || 0);
        const key = `${itemCode}_${rate.toFixed(4)}`;
        
        if (consolidatedMap[key]) {
          consolidatedMap[key].qty += qty;
          consolidatedMap[key].amount = consolidatedMap[key].qty * consolidatedMap[key].rate;
        } else {
          consolidatedMap[key] = {
            item_code: itemCode,
            qty: qty,
            rate: rate,
            amount: qty * rate,
            uom: 'Nos',
            expense_account: docType === 'Purchase Invoice' ? activeProfile.expense_account : undefined,
            cost_center: docType === 'Purchase Invoice' ? activeProfile.cost_center : undefined,
            schedule_date: docType === 'Purchase Order' ? (invoiceForm.date || new Date().toISOString().split('T')[0]) : undefined
          };
        }
      }
      let resolvedItems = Object.values(consolidatedMap);

      let taxesChildTable = [];
      const activeTaxTemplate = activeProfile.tax_template;
      let calculatedTaxTotal = 0;
      let netTotal = 0;
      let grandTotal = 0;

      let taxTemplateRows = [];
      let totalTaxRate = 0;

      if (activeTaxTemplate) {
        try {
          logToConsole(`Fetching Purchase Taxes and Charges Template "${activeTaxTemplate}" details from ERPNext...`, 'info');
          const taxTemplate = await erpRequest(`/api/resource/Purchase Taxes and Charges Template/${encodeURIComponent(activeTaxTemplate)}`);
          if (taxTemplate && taxTemplate.data && taxTemplate.data.taxes) {
            taxTemplateRows = taxTemplate.data.taxes;
            taxTemplateRows.forEach(row => {
               if (row.charge_type === "On Net Total" || row.charge_type === "On Previous Row Amount") {
                 totalTaxRate += parseFloat(row.rate || 0);
               }
            });
          }
        } catch (taxErr) {
          logToConsole(`Warning: Failed to fetch Tax template "${activeTaxTemplate}" rows: ${taxErr.message}. Syncing without explicit tax table.`, 'warn');
        }
      }

      // If tax inclusive, adjust resolved items back down to their base amounts
      if (taxInclusive && totalTaxRate > 0) {
        const divisor = 1 + (totalTaxRate / 100);
        resolvedItems = resolvedItems.map(item => {
          const newAmt = parseFloat(item.amount || 0) / divisor;
          const newRate = parseFloat(item.rate || 0) / divisor;
          return {
            ...item,
            amount: newAmt,
            rate: newRate,
            real_rate: item.real_rate ? (parseFloat(item.real_rate) / divisor) : undefined
          };
        });
        logToConsole(`s Tax Conclusive Mode: Divided item amounts by ${divisor.toFixed(4)} to extract base net total before applying taxes.`, 'info');
      }

      // Calculate Net Total of resolved items
      netTotal = resolvedItems.reduce((sum, item) => sum + parseFloat(item.amount || 0), 0) - (invoiceForm.discount_amount || 0);

      if (taxTemplateRows.length > 0) {
        let currentCumulative = netTotal;
        taxesChildTable = taxTemplateRows.map(row => {
          let taxAmt = 0;
          const rate = parseFloat(row.rate || 0);
          if (row.charge_type === "On Net Total") {
            taxAmt = netTotal * (rate / 100);
          } else if (row.charge_type === "On Previous Row Amount") {
            taxAmt = netTotal * (rate / 100); 
          } else {
            taxAmt = rate; 
          }
          
          taxAmt = Math.round(taxAmt * 100) / 100;
          currentCumulative += taxAmt;
          calculatedTaxTotal += taxAmt;
          
          return {
            charge_type: row.charge_type,
            account_head: row.account_head,
            rate: row.rate,
            description: row.description,
            category: row.category,
            add_deduct_tax: row.add_deduct_tax,
            cost_center: row.cost_center,
            included_in_print_rate: row.included_in_print_rate,
            tax_amount: taxAmt,
            base_tax_amount: taxAmt,
            total: currentCumulative,
            base_total: currentCumulative
          };
        });
        grandTotal = currentCumulative;
        logToConsole(`o" Successfully loaded Tax template "${activeTaxTemplate}" (${taxesChildTable.length} tax rows). Calculated Tax: $${calculatedTaxTotal.toFixed(2)}`, 'success');
      } else {
        grandTotal = netTotal;
      }
      
      let payload = {};
      let endpoint = '';
      
      if (docType === 'Purchase Order') {
        payload = {
          supplier: selectedSupplierObject.name,
          supplier_refernce: invoiceForm.invoice_number, // Spelled with typo for standard compatibility
          supplier_reference: invoiceForm.invoice_number, // Fallback for customized doctypes
          remarks: `Invoice Sync Automation. Supplier Invoice / Sales Order No: ${invoiceForm.invoice_number}.`, // Backup for resilient queries
          transaction_date: invoiceForm.date || new Date().toISOString().split('T')[0],
          schedule_date: invoiceForm.date || new Date().toISOString().split('T')[0],
          company: activeProfile.company,
          net_total: netTotal,
          base_net_total: netTotal,
          grand_total: grandTotal,
          base_grand_total: grandTotal,
          rounded_total: Math.round(grandTotal),
          base_rounded_total: Math.round(grandTotal),
          total_taxes_and_charges: calculatedTaxTotal,
          base_total_taxes_and_charges: calculatedTaxTotal,
          additional_discount_amount: invoiceForm.discount_amount || 0,
          discount_amount: invoiceForm.discount_amount || 0,
          apply_discount_on: invoiceForm.discount_amount > 0 ? "Net Total" : undefined,
          taxes_and_charges: activeTaxTemplate || undefined,
          taxes: activeTaxTemplate ? taxesChildTable : undefined,
          items: resolvedItems
        };
        // Dynamically inject references using discovered fields list
        if (poRefField) {
          payload[poRefField] = invoiceForm.invoice_number;
        }
        endpoint = '/api/resource/Purchase Order';
      } else {
        payload = {
          supplier: selectedSupplierObject.name,
          bill_no: invoiceForm.invoice_number,
          bill_date: invoiceForm.date || new Date().toISOString().split('T')[0],
          posting_date: invoiceForm.date || new Date().toISOString().split('T')[0],
          due_date: invoiceForm.date || new Date().toISOString().split('T')[0],
          payment_terms_template: "",
          payment_schedule: [],
          company: activeProfile.company,
          net_total: netTotal,
          base_net_total: netTotal,
          grand_total: grandTotal,
          base_grand_total: grandTotal,
          rounded_total: Math.round(grandTotal),
          base_rounded_total: Math.round(grandTotal),
          total_taxes_and_charges: calculatedTaxTotal,
          base_total_taxes_and_charges: calculatedTaxTotal,
          additional_discount_amount: invoiceForm.discount_amount || 0,
          discount_amount: invoiceForm.discount_amount || 0,
          apply_discount_on: invoiceForm.discount_amount > 0 ? "Net Total" : undefined,
          taxes_and_charges: activeTaxTemplate || undefined,
          taxes: activeTaxTemplate ? taxesChildTable : undefined,
          items: resolvedItems
        };
        endpoint = '/api/resource/Purchase Invoice';
      }
      
      const responseData = await erpRequest(endpoint, 'POST', payload);
      
      if (responseData.data) {
        const docName = responseData.data.name;
        logToConsole(`✓ Draft ${docType} ${docName} created successfully in ERPNext!`, 'success');
        
        let finalDocName = docName;
        let finalDocType = docType;
        
        if (docType === 'Purchase Order' && autoSubmitFlow) {
          try {
            // 1. Submit the Purchase Order
            logToConsole(`Submitting Purchase Order ${docName} in ERPNext...`, 'info');
            await erpRequest(`/api/resource/Purchase Order/${docName}`, 'PUT', { docstatus: 1 });
            logToConsole(`✓ Purchase Order ${docName} submitted successfully.`, 'success');
            
            // 2. Generate and Submit Purchase Receipt (PR)
            logToConsole(`Generating Purchase Receipt from Purchase Order ${docName}...`, 'info');
            const prTemplate = await erpRequest('/api/method/erpnext.buying.doctype.purchase_order.purchase_order.make_purchase_receipt', 'POST', { source_name: docName });
            if (prTemplate && prTemplate.message) {
              const prPayload = prTemplate.message;
              
              // Override warehouse if specified in activeProfile
              if (activeProfile.warehouse) {
                prPayload.items.forEach(item => {
                  item.warehouse = activeProfile.warehouse;
                });
              }
              
              // Map Serial Numbers if tracking is enabled
              if (activeProfile.enable_sn_tracking) {
                prPayload.items.forEach((item, idx) => {
                  // PR items are generated in the exact same order as PO items
                  const correspondingResolvedItem = resolvedItems[idx];
                  if (correspondingResolvedItem && correspondingResolvedItem.item_code === item.item_code && correspondingResolvedItem.serial_nos) {
                    // ERPNext expects serial numbers separated by newlines
                    const formattedSNs = correspondingResolvedItem.serial_nos
                      .split(/[\n,]+/) 
                      .map(s => s.trim())
                      .filter(s => s)
                      .join('\n');
                    item.serial_no = formattedSNs;
                  }
                });
              }
              
              // Set the transaction date to match the invoice date
              prPayload.posting_date = invoiceForm.date || new Date().toISOString().split('T')[0];
              prPayload.docstatus = 1; // submit automatically
              logToConsole("Submitting Purchase Receipt...", 'info');
              const prResult = await erpRequest('/api/resource/Purchase Receipt', 'POST', prPayload);
              if (prResult && prResult.data) {
                logToConsole(`✓ Purchase Receipt ${prResult.data.name} submitted successfully!`, 'success');
              }
            } else {
              throw new Error("Failed to retrieve Purchase Receipt template from ERPNext.");
            }
            
            // 3. Generate and Submit Purchase Invoice (PI)
            logToConsole(`Generating Purchase Invoice from Purchase Order ${docName}...`, 'info');
            const piTemplate = await erpRequest('/api/method/erpnext.buying.doctype.purchase_order.purchase_order.make_purchase_invoice', 'POST', { source_name: docName });
            if (piTemplate && piTemplate.message) {
              const piPayload = piTemplate.message;
              
              const invDate = invoiceForm.date || new Date().toISOString().split('T')[0];
              // Set the supplier invoice details and override dates to prevent future-date validations
              piPayload.bill_no = invoiceForm.invoice_number;
              piPayload.bill_date = invDate;
              piPayload.posting_date = invDate;
              piPayload.due_date = invDate;
              
              // Clear payment schedule and terms to prevent future due date validation blocks
              piPayload.payment_terms_template = "";
              piPayload.payment_schedule = [];
              
              piPayload.docstatus = 1; // submit automatically
              
              logToConsole("Submitting Purchase Invoice...", 'info');
              const piResult = await erpRequest('/api/resource/Purchase Invoice', 'POST', piPayload);
              if (piResult && piResult.data) {
                logToConsole(`✓ Purchase Invoice ${piResult.data.name} submitted successfully!`, 'success');
                // Point user review to the final invoice
                finalDocName = piResult.data.name;
                finalDocType = 'Purchase Invoice';
              }
            } else {
              throw new Error("Failed to retrieve Purchase Invoice template from ERPNext.");
            }
            
          } catch (autoErr) {
            logToConsole(`⚠ Automation Flow Error: ${autoErr.message}. PO was created and submitted, but receipt/invoice creation failed.`, 'warn');
          }
        }
        
        // Auto-save calculated rebate records to Neon Postgres (Only for VISKOU invoices)
        try {
          const supplierName = (selectedSupplierObject?.supplier_name || selectedSupplierObject?.name || invoiceForm?.supplier || supplierQuery || '').toLowerCase();
          const isViskou = supplierName.includes('viskou');

          if (!isViskou) {
            logToConsole(`ℹ Skipping rebate records: Rebate only applies to supplier VISKOU (current supplier: "${selectedSupplierObject?.supplier_name || invoiceForm.supplier || 'Non-VISKOU'}").`, 'info');
          } else {
            const rebateRecordsToSave = [];
            
            invoiceForm.items.forEach(item => {
              if (parseFloat(item.rate || 0) < 0) return; // Skip discount rows
              
              // Pick cleanest candidate model code (avoiding generic fallback 'Default Item')
              let modelCode = '';
              if (item.description && item.description !== 'Default Item' && item.description.trim().length > 0) {
                modelCode = item.description.trim();
              } else if (item.item_code && item.item_code !== 'Default Item' && item.item_code.trim().length > 0) {
                modelCode = item.item_code.trim();
              } else {
                modelCode = (item.description || item.item_code || 'ITEM').trim();
              }

              const upperCode = modelCode.toUpperCase();
              let rebatePct = 0;
              
              if (upperCode.startsWith('RG')) {
                rebatePct = 0.05;
              } else if (upperCode.startsWith('VS')) {
                rebatePct = 0.10;
              } else if (upperCode.startsWith('DS') || upperCode.startsWith('IDS') || upperCode.startsWith('CS')) {
                const isInException = rebateExceptionList.some(ex => {
                  const cleanEx = String(ex || '').toUpperCase().trim();
                  return cleanEx === upperCode || upperCode.includes(cleanEx) || cleanEx.includes(upperCode);
                });
                rebatePct = isInException ? 0.08 : 0.16;
              } else {
                rebatePct = 0;
              }
              
              const qty = parseFloat(item.qty || 1);
              const rate = parseFloat(item.rate || 0);
              const realRate = item.real_rate !== undefined && item.real_rate !== null ? parseFloat(item.real_rate) : rate;
              const rebateAmt = qty * realRate * rebatePct;
              
              if (rebateAmt > 0 || rebatePct > 0) {
                rebateRecordsToSave.push({
                  date: invoiceForm.date || new Date().toISOString().split('T')[0],
                  invoice_no: finalDocName || invoiceForm.invoice_number,
                  item_model_code: modelCode,
                  quantity: qty,
                  rate: rate,
                  real_rate: realRate,
                  rebate_percentage: rebatePct,
                  rebate_amount: rebateAmt,
                  company_profile: activeProfileName,
                  supplier_invoice_no: invoiceForm.invoice_number
                });
              }
            });

            if (rebateRecordsToSave.length > 0) {
              logToConsole(`Auto-saving ${rebateRecordsToSave.length} rebate record(s) to Neon DB for profile "${activeProfileName}"...`, 'info');
              const rebRes = await fetch('/api/rebates/records', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ records: rebateRecordsToSave })
              });
              const rebData = await rebRes.json();
              if (rebData.success) {
                logToConsole(`✓ ${rebData.message}`, 'success');
              }
            }
          }
        } catch (rebErr) {
          logToConsole(`⚠ Rebate storage warning: ${rebErr.message}`, 'warn');
        }

        logAction(
          "SYNC_SUCCESS",
          selectedSupplierObject.supplier_name,
          invoiceForm.grand_total,
          finalDocName,
          { docType: finalDocType, discount: invoiceForm.discount_amount, poName: docName }
        );
        
        const docUrl = `${activeProfile.url}/app/${finalDocType.toLowerCase().replace(' ', '-')}/${finalDocName}`;
        logToConsole(`👉 Review document: ${docUrl}`, 'info');
        
        alert(`Success! Created documents in ERPNext.`);
      }
      
    } catch (error) {
      logToConsole(`✗ Sync Failed: ${error.message}`, 'err');
      alert(`Sync Failed: ${error.message}`);
    }
  }

  // Live Rebate Item Evaluator
  function getItemRebateInfo(item) {
    if (parseFloat(item.rate || 0) < 0) return { pct: 0, amt: 0, label: 'Discount Row' };

    // Rebates strictly apply only to supplier VISKOU invoices
    const currentSupplier = (selectedSupplierObject?.supplier_name || selectedSupplierObject?.name || invoiceForm?.supplier || supplierQuery || '').toLowerCase();
    if (!currentSupplier.includes('viskou')) {
      return { pct: 0, amt: 0, label: 'Non-VISKOU' };
    }

    let modelCode = '';
    if (item.description && item.description !== 'Default Item' && item.description.trim().length > 0) {
      modelCode = item.description.trim();
    } else if (item.item_code && item.item_code !== 'Default Item' && item.item_code.trim().length > 0) {
      modelCode = item.item_code.trim();
    } else {
      modelCode = (item.description || item.item_code || '').trim();
    }

    const upperCode = modelCode.toUpperCase();
    let pct = 0;
    let label = '';

    if (upperCode.startsWith('RG')) {
      pct = 0.05;
      label = '5% (RG)';
    } else if (upperCode.startsWith('VS')) {
      pct = 0.10;
      label = '10% (VS)';
    } else if (upperCode.startsWith('DS') || upperCode.startsWith('IDS') || upperCode.startsWith('CS')) {
      const isInMasterlist = rebateExceptionList.some(ex => {
        const cleanEx = String(ex || '').toUpperCase().trim();
        return cleanEx === upperCode || upperCode.includes(cleanEx) || cleanEx.includes(upperCode);
      });
      if (isInMasterlist) {
        pct = 0.08;
        label = '8% (Masterlist)';
      } else {
        pct = 0.16;
        label = '16% (DS/iDS/CS Standard)';
      }
    } else {
      pct = 0;
      label = '0%';
    }

    const qty = parseFloat(item.qty || 1);
    const realRate = item.real_rate !== undefined && item.real_rate !== null ? parseFloat(item.real_rate) : parseFloat(item.rate || 0);
    const amt = qty * realRate * pct;

    return { pct, amt, label };
  }

  // Rebate & Real Rate Calculation Helpers
  function computeLineItemRealRates(chatPricesStr, items) {
    const prices = String(chatPricesStr || '')
      .split(',')
      .map(p => parseFloat(p.trim()))
      .filter(p => !isNaN(p));

    let physicalIdx = 0;
    return items.map(item => {
      const isDiscountRow = parseFloat(item.rate || 0) < 0 || (item.description || '').toLowerCase().includes('discount');
      if (isDiscountRow) {
        return { ...item, real_rate: item.rate };
      }
      // If index is within Chat Prices array length, use chat price; otherwise default to item rate
      const assignedRealRate = physicalIdx < prices.length ? prices[physicalIdx] : parseFloat(item.rate || 0);
      physicalIdx++;
      return { ...item, real_rate: assignedRealRate };
    });
  }

  function computeExpectedDiscount(items) {
    let expected = 0;
    (items || []).forEach(it => {
      if (parseFloat(it.rate || 0) >= 0) {
        const qty = parseFloat(it.qty || 1);
        const rate = parseFloat(it.rate || 0);
        const realRate = it.real_rate !== undefined && it.real_rate !== null ? parseFloat(it.real_rate) : rate;
        expected += (qty * rate) - (qty * realRate);
      }
    });
    return Math.max(0, expected);
  }

  // Item Autocomplete matching algorithm
  function findMatchingItemCode(desc) {
    const defaultFallback = desc ? desc.trim() : (parserSettings.defaultItem || '');
    if (!desc || itemsList.length === 0) return { item_code: defaultFallback, match_confidence: 'none' };
    
    // Sanitize query: standardize dashes, zero-width spaces, trim spaces and leading/trailing punctuation
    const query = desc
      .toLowerCase()
      .replace(/[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]/g, '-') // standardizes dashes
      .replace(/[\u00a0\u200b\u200c\u200d]/g, ' ')                  // standardizes zero-width spaces
      .trim()
      .replace(/^[\s.,:;'"!?,]+|[\s.,:;'"!?,]+$/g, '');             // trims trailing punctuation
    
    if (!query) {
      const fb = itemsList[0]?.item_code || itemsList[0]?.name || defaultFallback;
      return { item_code: fb, match_confidence: 'none' };
    }
    
    // Levenshtein similarity calculator helper
    function getLevenshteinSimilarity(s1, s2) {
      if (!s1 || !s2) return 0;
      const len1 = s1.length;
      const len2 = s2.length;
      
      const matrix = [];
      for (let i = 0; i <= len1; i++) {
        matrix[i] = [i];
      }
      for (let j = 0; j <= len2; j++) {
        matrix[0][j] = j;
      }
      
      for (let i = 1; i <= len1; i++) {
        for (let j = 1; j <= len2; j++) {
          const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
          matrix[i][j] = Math.min(
            matrix[i - 1][j] + 1,      // deletion
            matrix[i][j - 1] + 1,      // insertion
            matrix[i - 1][j - 1] + cost // substitution
          );
        }
      }
      
      const distance = matrix[len1][len2];
      const maxLen = Math.max(len1, len2);
      return maxLen === 0 ? 1 : 1 - (distance / maxLen);
    }

    function getItemCode(i) {
      return (i.item_code || i.name || '').toLowerCase().trim();
    }

    function getItemName(i) {
      return (i.item_name || '').toLowerCase().trim();
    }

    // 1. Exact match on code (highest priority)
    const exactCode = itemsList.find(i => getItemCode(i) === query);
    if (exactCode) return { item_code: exactCode.item_code || exactCode.name, match_confidence: 'high' };
    
    // 2. Exact match on name (secondary priority)
    const exactName = itemsList.find(i => getItemName(i) === query);
    if (exactName) return { item_code: exactName.item_code || exactName.name, match_confidence: 'high' };
    
    // 3. Score all candidates
    const candidates = [];
    
    itemsList.forEach(i => {
      const code = getItemCode(i);
      const name = getItemName(i);
      
      // -- Code matching metrics
      let codeSubScore = 0;
      if (query.includes(code)) codeSubScore = code.length;
      else if (code.includes(query)) codeSubScore = query.length;
      
      let codeTokenMatches = 0;
      const codeTokens = code.split(/[\s,._/#-]+/).filter(t => t.length >= 2);
      codeTokens.forEach(token => { if (query.includes(token)) codeTokenMatches++; });
      const codeTokenRatio = codeTokens.length > 0 ? (codeTokenMatches / codeTokens.length) : 0;
      const codeLev = getLevenshteinSimilarity(query, code);
      const codeScore = codeSubScore + (codeTokenRatio * 15.0) + (codeLev * 15.0);
      
      // -- Name matching metrics
      let nameSubScore = 0;
      if (query.includes(name)) nameSubScore = name.length;
      else if (name.includes(query)) nameSubScore = query.length;
      
      let nameTokenMatches = 0;
      const nameTokens = name.split(/[\s,._/#-]+/).filter(t => t.length >= 2);
      nameTokens.forEach(token => { if (query.includes(token)) nameTokenMatches++; });
      const nameTokenRatio = nameTokens.length > 0 ? (nameTokenMatches / nameTokens.length) : 0;
      const nameLev = getLevenshteinSimilarity(query, name);
      const nameScore = nameSubScore + (nameTokenRatio * 10.0) + (nameLev * 10.0);
      
      candidates.push({ item: i, codeScore, nameScore });
    });
    
    candidates.sort((a, b) => {
      if (Math.abs(a.codeScore - b.codeScore) > 0.001) return b.codeScore - a.codeScore;
      return b.nameScore - a.nameScore;
    });
    
    const best = candidates[0];
    if (best) {
      if (best.codeScore > 4.0 || best.nameScore > 4.0) {
         return { item_code: best.item.item_code || best.item.name, match_confidence: 'high' };
      } else if (best.codeScore > 1.5 || best.nameScore > 1.5) {
         return { item_code: best.item.item_code || best.item.name, match_confidence: 'low' };
      }
    }
    
    return { item_code: defaultFallback, match_confidence: 'none' };
  }
  async function loadUnpaidInvoices(supplierId) {
    setLoadingInvoices(true);
    setUnpaidInvoices([]);
    setCheckedInvoiceIds({});
    setMatchedInvoiceSequence([]);
    setPaymentAmount(0);
    setPaymentListStatus("Loading unpaid invoices from ERPNext...");
    setPaymentListColor("var(--status-testing)");
    try {
      const filterStr = JSON.stringify([
        ["supplier", "=", supplierId],
        ["docstatus", "=", 1],
        ["outstanding_amount", ">", 0],
        ["status", "!=", "Paid"]
      ]);
      const fieldsStr = JSON.stringify(["name", "posting_date", "grand_total", "outstanding_amount", "bill_no", "status"]);
      const res = await erpRequest(`/api/resource/Purchase Invoice?filters=${encodeURIComponent(filterStr)}&fields=${encodeURIComponent(fieldsStr)}&order_by=posting_date%20desc&limit_page_length=1000`);
      
      const invoices = (res.data || []).filter(inv => 
        inv.outstanding_amount > 0.01 && 
        inv.status !== 'Paid' && 
        inv.status !== 'Completed'
      );
      setUnpaidInvoices(invoices);
      
      if (invoices.length === 0) {
        setPaymentListStatus(`No unpaid invoices found for ${supplierId}.`);
        setPaymentListColor("var(--status-success)");
      } else {
        const targetNo = (invoiceForm.invoice_number || '').trim().toLowerCase();
        if (targetNo) {
          const matched = invoices.filter(inv => {
            const bNo = (inv.bill_no || '').trim().toLowerCase();
            const nameVal = (inv.name || '').trim().toLowerCase();
            return bNo === targetNo || nameVal === targetNo;
          });
          
          if (matched.length > 0) {
            const targetAmount = parseFloat(invoiceForm.grand_total || 0);
            const sameAmountInvoices = matched.filter(inv => Math.abs(parseFloat(inv.grand_total || inv.outstanding_amount || 0) - targetAmount) < 0.02);
            
            const autoChecked = {};
            const autoSeq = [];
            let autoSum = 0;
            const checkTargets = sameAmountInvoices.length > 0 ? sameAmountInvoices : matched;
            checkTargets.forEach(m => {
              autoChecked[m.name] = true;
              autoSeq.push(m.name);
              autoSum += m.outstanding_amount;
            });
            setCheckedInvoiceIds(autoChecked);
            setMatchedInvoiceSequence(autoSeq);
            setPaymentAmount(autoSum);

            if (invoiceForm.invoice_number) {
              setPaymentRefNo(invoiceForm.invoice_number);
            }
            if (invoiceForm.date) {
              setPaymentRefDate(invoiceForm.date);
            }
            
            if (sameAmountInvoices.length > 0) {
              setPaymentListStatus(`Showing ${sameAmountInvoices.length} matching invoice(s) with same amount of $${targetAmount.toFixed(2)}.`);
              setPaymentListColor("var(--status-success)");
            } else {
              setPaymentListStatus(`⚠️ Found matching reference but amount does not match $${targetAmount.toFixed(2)}.`);
              setPaymentListColor("var(--color-secondary)");
            }
          } else {
            setPaymentListStatus(`⚠️ No unpaid invoice matching parsed reference "${invoiceForm.invoice_number}" was found.`);
            setPaymentListColor("var(--color-secondary)");
          }
        } else {
          setPaymentListStatus(`Loaded ${invoices.length} unpaid invoices.`);
          setPaymentListColor("var(--text-muted)");
        }
      }
    } catch (e) {
      setPaymentListStatus(`Failed to load: ${e.message}`);
      setPaymentListColor("var(--status-error)");
    } finally {
      setLoadingInvoices(false);
    }
  }

  // Recalculate checked invoices total & sync sequence
  const handleInvoiceCheckChange = (invId, outstanding) => {
    setCheckedInvoiceIds(prev => {
      const isChecking = !prev[invId];
      const next = { ...prev, [invId]: isChecking };
      let sum = 0;
      unpaidInvoices.forEach(inv => {
        if (next[inv.name]) {
          sum += inv.outstanding_amount;
        }
      });
      setPaymentAmount(sum);
      
      setMatchedInvoiceSequence(prevSeq => {
        if (isChecking) {
          if (!prevSeq.includes(invId)) return [...prevSeq, invId];
          return prevSeq;
        } else {
          return prevSeq.filter(id => id !== invId);
        }
      });

      return next;
    });
  };

  // Duplicate Check logic
  function calculateGlobalDuplicates(invoices) {
    const amountCounts = {};
    const billNoCounts = {};
    invoices.forEach(inv => {
      const amt = inv.outstanding_amount.toFixed(2);
      const bill = (inv.bill_no || '').trim().toLowerCase();
      amountCounts[amt] = (amountCounts[amt] || 0) + 1;
      if (bill && bill !== 'n/a') {
        billNoCounts[bill] = (billNoCounts[bill] || 0) + 1;
      }
    });
    return { amountCounts, billNoCounts };
  }

  // Compute duplicate parameters globally for all loaded invoices
  const globalDuplicates = calculateGlobalDuplicates(unpaidInvoices);

  // Filter and sort unpaid invoices array
  const getProcessedInvoicesList = () => {
    const query = payInvoiceSearch.trim().toLowerCase();
    
    let filtered = unpaidInvoices;
    
    // Apply search query filter if user typed anything
    if (query) {
      filtered = filtered.filter(inv => {
        const invId = inv.name.toLowerCase();
        const billNo = (inv.bill_no || '').toLowerCase();
        const amtStr = inv.outstanding_amount.toFixed(2);
        const amtClean = inv.outstanding_amount.toString();
        
        return invId.includes(query) || billNo.includes(query) || amtStr.includes(query) || amtClean.includes(query);
      });
    }

    // Helper to extract the unique group key for duplicates
    const getSortKey = (inv) => {
      const isDupAmt = globalDuplicates.amountCounts[inv.outstanding_amount.toFixed(2)] > 1;
      if (isDupAmt) return `amt_${inv.outstanding_amount.toFixed(2)}`;
      
      const bill = (inv.bill_no || '').trim().toLowerCase();
      const isDupBill = bill && bill !== 'n/a' && globalDuplicates.billNoCounts[bill] > 1;
      if (isDupBill) return `bill_${bill}`;
      
      return `uniq_${inv.name}`;
    };

    // Pre-calculate the max date for each group key to sort duplicate sets chronologically
    const groupMaxDates = {};
    filtered.forEach(inv => {
      const key = getSortKey(inv);
      const time = new Date(inv.posting_date || 0).getTime();
      groupMaxDates[key] = Math.max(groupMaxDates[key] || 0, time);
    });

    // Sort the list
    filtered.sort((a, b) => {
      const aChecked = checkedInvoiceIds[a.name] || false;
      const bChecked = checkedInvoiceIds[b.name] || false;
      
      // 1. Checked ones first
      if (aChecked !== bChecked) return aChecked ? -1 : 1;

      // 1b. If both are checked, sort in the exact sequence of the payment list document
      if (aChecked && bChecked) {
        const indexA = matchedInvoiceSequence.indexOf(a.name);
        const indexB = matchedInvoiceSequence.indexOf(b.name);
        if (indexA !== -1 && indexB !== -1) return indexA - indexB;
        if (indexA !== -1) return -1;
        if (indexB !== -1) return 1;
      }

      const keyA = getSortKey(a);
      const keyB = getSortKey(b);

      const aIsDup = !keyA.startsWith('uniq_');
      const bIsDup = !keyB.startsWith('uniq_');

      // 2. Duplicates first
      if (aIsDup !== bIsDup) return aIsDup ? -1 : 1;

      // 3. Sort duplicate groups by their newest date
      if (keyA !== keyB) {
        const dateDiff = (groupMaxDates[keyB] || 0) - (groupMaxDates[keyA] || 0);
        if (dateDiff !== 0) return dateDiff;
        return keyA.localeCompare(keyB);
      }

      // 4. Same group: sort newest first
      return new Date(b.posting_date || 0) - new Date(a.posting_date || 0);
    });

    return filtered;
  };

  // Payment List File matching
  const triggerPayListSelect = () => payListInputRef.current.click();

  const handlePayListFileChange = async (e) => {
    if (e.target.files.length === 0) return;
    const file = e.target.files[0];
    
    setPaymentListStatus("Processing payment list file...");
    setPaymentListColor("var(--status-testing)");
    logToConsole(`Uploading payment list file: ${file.name} (${Math.round(file.size / 1024)} KB)...`, 'info');
    
    if (file.type.startsWith('image/')) {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        await extractInvoiceCodesFromImage(evt.target.result);
      };
      reader.readAsDataURL(file);
    } else if (file.type === 'application/pdf') {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        try {
          const arrayBuffer = evt.target.result;
          
          // Dynamically import PDF.js client-side
          const pdfjsLib = await import('pdfjs-dist');
          pdfjsLib.GlobalWorkerOptions.workerSrc = `//unpkg.com/pdfjs-dist@6.1.200/build/pdf.worker.min.mjs`;

          const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
          const pdf = await loadingTask.promise;
          
          let text = '';
          for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            text += textContent.items.map(item => item.str).join(' ') + '\n';
          }
          
          if (text.trim().length > 10) {
            await extractInvoiceCodesFromText(text);
          } else {
            // Scanned PDF
            const page = await pdf.getPage(1);
            const scale = 2.0;
            const viewport = page.getViewport({ scale });
            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d');
            canvas.height = viewport.height;
            canvas.width = viewport.width;
            await page.render({ canvasContext: context, viewport }).promise;
            await extractInvoiceCodesFromImage(canvas.toDataURL('image/png'));
          }
        } catch (err) {
          setPaymentListStatus(`PDF Reading failed: ${err.message}`);
          setPaymentListColor("var(--status-error)");
        }
      };
      reader.readAsArrayBuffer(file);
    } else if (file.type === 'text/plain' || file.name.endsWith('.txt')) {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        await extractInvoiceCodesFromText(evt.target.result);
      };
      reader.readAsText(file);
    } else {
      setPaymentListStatus("Unsupported format. Select text, PDF, or image.");
      setPaymentListColor("var(--status-error)");
    }
  };

  // Robust Cloud AI fetch helper with automatic 3-attempt retry loop & safety settings
  async function callCloudAIWithRetry(fetchFn, maxRetries = 3) {
    let lastError = null;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const response = await fetchFn();
        const resData = await response.json();

        // 1. Check top-level API error response
        if (resData && resData.error) {
          const errMsg = resData.error.message || JSON.stringify(resData.error);
          throw new Error(`Cloud AI Error (${resData.error.code || response.status}): ${errMsg}`);
        }
        
        // 2. Check Gemini safety block
        if (resData && resData.promptFeedback && resData.promptFeedback.blockReason) {
          throw new Error(`Gemini Blocked Prompt: ${resData.promptFeedback.blockReason}`);
        }

        // 3. Extract Gemini response candidates safely
        if (resData && resData.candidates && resData.candidates.length > 0) {
          const candidate = resData.candidates[0];
          if (candidate.finishReason === "SAFETY") {
            throw new Error("Gemini safety filter blocked the content.");
          }
          if (candidate.content && candidate.content.parts && candidate.content.parts.length > 0) {
            const rawText = candidate.content.parts[0].text || '';
            if (rawText.trim()) {
              return JSON.parse(rawText.trim());
            }
          }
        }

        // 4. Extract OpenAI choices safely
        if (resData && resData.choices && resData.choices.length > 0) {
          const content = resData.choices[0].message?.content;
          if (content && content.trim()) {
            return JSON.parse(content.trim());
          }
        }

        throw new Error(`Cloud AI returned HTTP ${response.status} with empty response structure.`);
      } catch (err) {
        lastError = err;
        logToConsole(`Cloud AI Attempt ${attempt}/${maxRetries} failed: ${err.message}`, attempt < maxRetries ? 'warn' : 'err');
        if (attempt < maxRetries) {
          await new Promise(r => setTimeout(r, attempt * 800));
        }
      }
    }
    throw lastError;
  }

  async function extractInvoiceCodesFromImage(dataUrl) {
    const engineType = parserSettings.type;
    
    if (engineType === 'regex') {
      try {
        setPaymentListStatus("Running local OCR scan...");
        
        // Dynamically import tesseract client-side
        const { createWorker } = await import('tesseract.js');
        const worker = await createWorker('eng');
        const result = await worker.recognize(dataUrl);
        await worker.terminate();
        await extractInvoiceCodesFromText(result.data.text);
      } catch (err) {
        setPaymentListStatus(`Local OCR failed: ${err.message}`);
        setPaymentListColor("var(--status-error)");
      }
      return;
    }
    
    try {
      setPaymentListStatus("Calling Cloud AI for image matching...");
      const base64Parts = dataUrl.split(',');
      const mimeType = base64Parts[0].match(/:(.*?);/)[1];
      const base64Data = base64Parts[1];
      
      let extractedResult = null;
      if (engineType === 'gemini') {
        if (!parserSettings.geminiKey) throw new Error("Gemini key missing.");
        
        extractedResult = await callCloudAIWithRetry(() => 
          fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${parserSettings.geminiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{
                parts: [
                  { text: "Analyze this payment list document. Identify the Supplier/Vendor Name (supplier), and extract all listed Tax Invoice Numbers, Bill Numbers, SO Numbers, reference codes, and line amounts (or debits). IMPORTANT: Do NOT extract the running balance or cumulative total as the amount. Extract ONLY the individual line amount or debit for that specific row. Strip trailing hyphens or dashes from invoice numbers. If a line contains both an SO number and a Tax Invoice number, extract BOTH into their respective fields. Respond ONLY with a valid JSON object matching this schema: {\"supplier\":\"Supplier Name or empty string\",\"items\":[{\"so_no\":\"SO251103022\",\"tax_invoice\":\"ST251103022\",\"amount\":32.70}]}" },
                  { inlineData: { mimeType, data: base64Data } }
                ]
              }],
              safetySettings: [
                { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
                { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
                { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
                { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
              ],
              generationConfig: { responseMimeType: 'application/json' }
            })
          })
        );
      } else if (engineType === 'openai') {
        if (!parserSettings.openaiKey) throw new Error("OpenAI key missing.");
        
        extractedResult = await callCloudAIWithRetry(() =>
          fetch("https://api.openai.com/v1/chat/completions", {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${parserSettings.openaiKey}`
            },
            body: JSON.stringify({
              model: parserSettings.openaiModel,
              response_format: { type: 'json_object' },
              messages: [
                { role: 'system', content: "Respond ONLY with a JSON object: { \"supplier\": \"Supplier Name\", \"items\": [{\"code\": \"ST251103022\", \"amount\": 32.70}] }. Strip trailing dashes or hyphens from invoice codes." },
                { role: 'user', content: [
                  { type: 'text', text: "Analyze payment list. Extract supplier name and invoice codes/amounts without trailing dashes." },
                  { type: 'image_url', image_url: { url: dataUrl } }
                ]}
              ]
            })
          })
        );
      }
      
      await matchAndCheckInvoices(extractedResult);
    } catch (err) {
      setPaymentListStatus(`Cloud AI match failed: ${err.message}`);
      setPaymentListColor("var(--status-error)");
    }
  }

  async function extractInvoiceCodesFromText(text) {
    const engineType = parserSettings.type;
    
    if (engineType === 'regex') {
      const codeRegex = /\b[A-Za-z0-9#\/-]{3,25}\b/g;
      const amountRegex = /\b\d+(?:\.\d{2})\b/g;
      const matchesCode = text.match(codeRegex) || [];
      const matchesAmt = text.match(amountRegex) || [];
      await matchAndCheckInvoices([...new Set([...matchesCode, ...matchesAmt])]);
      return;
    }
    
    try {
      setPaymentListStatus("Calling Cloud AI for text matching...");
      let extractedResult = null;
      
      if (engineType === 'gemini') {
        if (!parserSettings.geminiKey) throw new Error("Gemini Key missing.");
        
        extractedResult = await callCloudAIWithRetry(() =>
          fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${parserSettings.geminiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{
                parts: [{
                  text: `Analyze this payment list text. Identify the Supplier/Vendor Name (supplier), and extract all listed Tax Invoice Numbers, Bill Numbers, SO Numbers, reference codes, and line amounts (or debits). IMPORTANT: Do NOT extract the running balance or cumulative total as the amount. Extract ONLY the individual line amount or debit for that specific row. Strip trailing hyphens or dashes from invoice numbers. If a line contains both an SO number and a Tax Invoice number, extract BOTH into their respective fields. Respond ONLY with a valid JSON object matching this schema: {\"supplier\":\"Supplier Name or empty string\",\"items\":[{\"so_no\":\"SO251103022\",\"tax_invoice\":\"ST251103022\",\"amount\":32.70}]}\n\nPayment List Text:\n${text}`
                }]
              }],
              safetySettings: [
                { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
                { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
                { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
                { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
              ],
              generationConfig: { responseMimeType: 'application/json' }
            })
          })
        );
      } else if (engineType === 'openai') {
        if (!parserSettings.openaiKey) throw new Error("OpenAI Key missing.");
        
        extractedResult = await callCloudAIWithRetry(() =>
          fetch("https://api.openai.com/v1/chat/completions", {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${parserSettings.openaiKey}`
            },
            body: JSON.stringify({
              model: parserSettings.openaiModel,
              response_format: { type: 'json_object' },
              messages: [
                { role: 'system', content: "Respond ONLY with a JSON object: { \"supplier\": \"Supplier Name\", \"items\": [{\"code\": \"ST251103022\", \"amount\": 32.70}] }" },
                { role: 'user', content: `Extract supplier name and invoice codes/amounts without trailing hyphens:\n${text}` }
              ]
            })
          })
        );
      }
      
      await matchAndCheckInvoices(extractedResult);
    } catch (err) {
      setPaymentListStatus(`Cloud AI match failed: ${err.message}`);
      setPaymentListColor("var(--status-error)");
    }
  }

  async function matchAndCheckInvoices(extractedData) {
    let detectedSupplier = "";
    let rawItems = [];

    if (typeof extractedData === 'object' && extractedData !== null) {
      if (extractedData.supplier) detectedSupplier = String(extractedData.supplier).trim();
      if (Array.isArray(extractedData.items)) {
        rawItems = extractedData.items;
      } else if (Array.isArray(extractedData.codes)) {
        rawItems = extractedData.codes;
      } else if (Array.isArray(extractedData)) {
        rawItems = extractedData;
      } else {
        rawItems = Object.values(extractedData);
      }
    } else if (Array.isArray(extractedData)) {
      rawItems = extractedData;
    }

    logToConsole(`Payment List Extracted Supplier: "${detectedSupplier}", Items count: ${rawItems.length}`, 'info');

    let currentUnpaid = unpaidInvoices;

    // Auto-match and select Supplier from document if detected or if user selected one
    if (detectedSupplier && suppliersList.length > 0) {
      const q = detectedSupplier.toLowerCase();
      let bestMatch = null;
      let maxScore = 0;
      
      suppliersList.forEach(s => {
        const sName = (s.supplier_name || s.name || "").toLowerCase();
        if (sName.includes(q) || q.includes(sName)) {
          const score = Math.max(sName.length, q.length);
          if (score > maxScore) {
            maxScore = score;
            bestMatch = s;
          }
        }
      });

      if (bestMatch) {
        logToConsole(`✓ Auto-matched Supplier from payment list: "${bestMatch.supplier_name || bestMatch.name}" (${bestMatch.name})`, 'success');
        setPaySelectedSupplier(bestMatch);
        setPaySupplierQuery(bestMatch.supplier_name || bestMatch.name || "");
        
        // Fetch fresh unpaid invoices for this supplier directly
        const freshInvoices = await loadUnpaidInvoices(bestMatch.name);
        if (freshInvoices && freshInvoices.length > 0) {
          currentUnpaid = freshInvoices;
        }
      }
    } else if (paySelectedSupplier) {
      const freshInvoices = await loadUnpaidInvoices(paySelectedSupplier.name);
      if (freshInvoices && freshInvoices.length > 0) {
        currentUnpaid = freshInvoices;
      }
    }

    const activeSupplierName = (paySelectedSupplier?.supplier_name || paySelectedSupplier?.name || detectedSupplier || '').toLowerCase();
    const isViskouSupplier = activeSupplierName.includes('viskou');

    // Filter out common noise tokens (years, generic words, pure numbers)
    const currentYear = new Date().getFullYear().toString();
    const noiseWords = new Set([currentYear, "2024", "2025", "2026", "2027", "page", "total", "subtotal", "date", "amount", "invoice", "bill", "payment", "list", "no", "num", "ref", "supplier", "singapore", "pte", "ltd"]);

    const codeTokens = [];
    const itemPairs = [];
    rawItems.forEach(it => {
      if (typeof it === 'object' && it !== null) {
        let amtRaw = String(it.amount || it.grand_total || it.outstanding_amount || it.debit || 0);
        // Remove commas from numbers (e.g. "1,058.39" -> "1058.39")
        amtRaw = amtRaw.replace(/,/g, '');
        const amtNum = parseFloat(amtRaw);
        
        // Extract all candidate string properties (code, tax_invoice, bill_no, so_no, reference, etc.)
        const candidateStrings = [];
        Object.keys(it).forEach(k => {
          if (k !== 'amount' && k !== 'grand_total' && k !== 'outstanding_amount' && k !== 'debit' && k !== 'credit' && k !== 'balance' && k !== 'date') {
            const val = String(it[k] || '').trim();
            if (val) candidateStrings.push(val);
          }
        });

        candidateStrings.forEach(codeStr => {
          // For Viskou or trailing punctuation, strip trailing dashes while preserving internal dashes for other suppliers
          let cleanCode = codeStr.trim();
          if (isViskouSupplier || /[-_/\s]+$/.test(cleanCode)) {
            cleanCode = cleanCode.replace(/[-_/\s]+$/, '').trim();
          }
          const norm = cleanCode.toLowerCase().replace(/[^a-z0-9]/g, '');
          
          if (norm.length >= 3 && !noiseWords.has(norm)) {
            codeTokens.push({ original: cleanCode.toLowerCase(), norm });
            if (!isNaN(amtNum) && amtNum > 0) {
              itemPairs.push({ codeNorm: norm, original: cleanCode.toLowerCase(), amount: amtNum });
            }
          }
        });
      } else if (it !== null && it !== undefined) {
        const str = String(it).trim();
        let cleanStr = str;
        if (isViskouSupplier || /[-_/\s]+$/.test(cleanStr)) {
          cleanStr = cleanStr.replace(/[-_/\s]+$/, '').trim();
        }
        const original = cleanStr.toLowerCase();
        const norm = original.replace(/[^a-z0-9]/g, '');
        
        // Skip pure numbers (amounts/years/page numbers) from being standalone code tokens
        const isPureNumber = /^\d+(\.\d+)?$/.test(cleanStr);
        if (!isPureNumber && norm.length >= 3 && !noiseWords.has(norm)) {
          codeTokens.push({ original, norm });
        }
      }
    });

    if (codeTokens.length === 0 && itemPairs.length === 0) {
      setPaymentListStatus("No valid invoice numbers or bill codes found in document.");
      setPaymentListColor("var(--status-warn)");
      if (payListInputRef.current) payListInputRef.current.value = '';
      return;
    }
    
    logToConsole(`Matching ${codeTokens.length} code token(s) against ${currentUnpaid.length} unpaid invoices...`, 'info');
    setPaymentListCodes(codeTokens.map(t => t.original));
    
    const nextChecks = {};
    const matchedSeq = [];
    let matchedCount = 0;
    let sum = 0;

    // Process tokens in exact document order to preserve payment list sequence
    codeTokens.forEach(tok => {
      const isNumeric = /^\d+$/.test(tok.norm);
      if (isNumeric && tok.norm.length < 4) return;

      currentUnpaid.forEach(inv => {
        if (nextChecks[inv.name]) return;
        const invId = inv.name.toLowerCase();
        const invIdNorm = invId.replace(/[^a-z0-9]/g, '');
        const billNo = (inv.bill_no || '').toLowerCase();
        const billNoNorm = billNo.replace(/[^a-z0-9]/g, '');

        let match = false;
        if (tok.norm.length >= 3) {
          if (invIdNorm.includes(tok.norm) || tok.norm.includes(invIdNorm)) match = true;
          if (billNoNorm && (billNoNorm.includes(tok.norm) || tok.norm.includes(billNoNorm))) match = true;
        }
        if (tok.original.length >= 4) {
          if (invId.includes(tok.original) || tok.original.includes(invId)) match = true;
          if (billNo && (billNo.includes(tok.original) || tok.original.includes(billNo))) match = true;
        }

        if (match) {
          nextChecks[inv.name] = true;
          matchedSeq.push(inv.name);
          matchedCount++;
          sum += inv.outstanding_amount;
        }
      });
    });

    itemPairs.forEach(pair => {
      currentUnpaid.forEach(inv => {
        if (nextChecks[inv.name]) return;
        const invIdNorm = inv.name.toLowerCase().replace(/[^a-z0-9]/g, '');
        const billNoNorm = (inv.bill_no || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        const outAmt = inv.outstanding_amount;
        const grandAmt = inv.grand_total;

        const cMatch = pair.codeNorm.length >= 3 && (invIdNorm.includes(pair.codeNorm) || (billNoNorm && billNoNorm.includes(pair.codeNorm)));
        // Relaxed amount matching per user request: only match the SO / Tax Invoice code
        
        if (cMatch) {
          nextChecks[inv.name] = true;
          matchedSeq.push(inv.name);
          matchedCount++;
          sum += inv.outstanding_amount;
        }
      });
    });

    setCheckedInvoiceIds(nextChecks);
    setMatchedInvoiceSequence(matchedSeq);
    setPaymentAmount(sum);
    
    if (matchedCount > 0) {
      setPaymentListStatus(`✓ Auto-checked ${matchedCount} matching invoice(s)!`);
      setPaymentListColor("var(--status-success)");
      logToConsole(`✓ Auto-checked ${matchedCount} invoice(s) ($${sum.toFixed(2)}) based on your payment list file.`, 'success');
      alert(`Successfully matched and checked ${matchedCount} invoice(s) ($${sum.toFixed(2)}) in the list.`);
    } else {
      setPaymentListStatus(`⚠️ All invoices in the document appear to be already paid or not found.`);
      setPaymentListColor("var(--status-warn)");
      logToConsole(`⚠️ None of the extracted codes matched open unpaid invoices.`, 'warn');
      alert(`None of the invoice codes in the document matched open unpaid invoices in your system. They may already be fully paid!`);
    }
    
    if (payListInputRef.current) payListInputRef.current.value = '';
  }

  // Create ERPNext Payment Entry
  async function submitPaymentEntry() {
    if (!paySelectedSupplier) {
      alert("Please select a supplier first.");
      return;
    }
    
    const checkedSet = new Set(Object.keys(checkedInvoiceIds).filter(id => checkedInvoiceIds[id]));
    const selectedRefs = [];
    
    // Maintain exact payment list document sequence for Payment Entry creation
    matchedInvoiceSequence.forEach(id => {
      if (checkedSet.has(id)) {
        selectedRefs.push(id);
        checkedSet.delete(id);
      }
    });
    checkedSet.forEach(id => selectedRefs.push(id));

    if (selectedRefs.length === 0) {
      alert("Please check at least one invoice to pay.");
      return;
    }

    if (!activeProfile.payment_account || !activeProfile.creditors_account) {
      alert("Payment account or Creditors account is not configured.");
      return;
    }

    logToConsole(`Preparing Payment Entry of $${paymentAmount.toFixed(2)} for ${paySelectedSupplier.supplier_name}...`, 'info');

    const references = [];
    for (const id of selectedRefs) {
      try {
        logToConsole(`Fetching installment details for ${id} from ERPNext...`, 'info');
        const invDoc = await erpRequest(`/api/resource/Purchase Invoice/${encodeURIComponent(id)}`);
        
        if (invDoc && invDoc.data) {
          const doc = invDoc.data;
          
          // Check if there is a payment schedule
          if (doc.payment_schedule && doc.payment_schedule.length > 0) {
            let addedScheduleRow = false;
            doc.payment_schedule.forEach(termRow => {
              let termOutstanding = parseFloat(termRow.outstanding_amount || 0);
              if (!termRow.outstanding_amount && termRow.payment_amount) {
                termOutstanding = parseFloat(termRow.payment_amount) - parseFloat(termRow.paid_amount || 0);
              }
              
              if (termOutstanding > 0.01) {
                references.push({
                  reference_doctype: 'Purchase Invoice',
                  reference_name: id,
                  total_amount: doc.grand_total,
                  outstanding_amount: termOutstanding,
                  allocated_amount: termOutstanding,
                  payment_term: termRow.payment_term
                });
                addedScheduleRow = true;
              }
            });
            
            if (!addedScheduleRow) {
              references.push({
                reference_doctype: 'Purchase Invoice',
                reference_name: id,
                total_amount: doc.grand_total,
                outstanding_amount: doc.outstanding_amount,
                allocated_amount: doc.outstanding_amount
              });
            }
          } else {
            references.push({
              reference_doctype: 'Purchase Invoice',
              reference_name: id,
              total_amount: doc.grand_total,
              outstanding_amount: doc.outstanding_amount,
              allocated_amount: doc.outstanding_amount
            });
          }
        } else {
          throw new Error(`Failed to retrieve invoice details for ${id}`);
        }
      } catch (err) {
        logToConsole(`Warning: Failed to fetch full details for invoice ${id}: ${err.message}. Falling back to default reference.`, 'warn');
        const inv = unpaidInvoices.find(i => i.name === id);
        if (inv) {
          references.push({
            reference_doctype: 'Purchase Invoice',
            reference_name: id,
            total_amount: inv.grand_total,
            outstanding_amount: inv.outstanding_amount,
            allocated_amount: inv.outstanding_amount
          });
        }
      }
    }

    const paymentPayload = {
      doctype: 'Payment Entry',
      payment_type: paymentType,
      party_type: 'Supplier',
      party: paySelectedSupplier.name,
      posting_date: paymentDate,
      mode_of_payment: paymentMethod,
      paid_from: paymentType === 'Pay' ? activeProfile.payment_account : activeProfile.creditors_account,
      paid_to: paymentType === 'Pay' ? activeProfile.creditors_account : activeProfile.payment_account,
      paid_amount: paymentAmount,
      received_amount: paymentAmount,
      target_exchange_rate: 1,
      reference_no: paymentRefNo,
      reference_date: paymentRefDate,
      references: references
    };

    try {
      const responseData = await erpRequest('/api/resource/Payment Entry', 'POST', paymentPayload);
      if (responseData.data) {
        const docName = responseData.data.name;
        logToConsole(`✓ Payment Entry ${docName} created successfully in ERPNext!`, 'success');
        
        // Auto submit the Payment Entry
        try {
          logToConsole(`Submitting Payment Entry ${docName} in ERPNext...`, 'info');
          await erpRequest(`/api/resource/Payment Entry/${docName}`, 'PUT', { docstatus: 1 });
          logToConsole(`✓ Payment Entry ${docName} submitted successfully in ERPNext!`, 'success');
        } catch (subErr) {
          logToConsole(`⚠ Failed to auto-submit Payment Entry ${docName}: ${subErr.message}`, 'warn');
        }

        logAction(
          "PAYMENT_CREATED",
          paySelectedSupplier.supplier_name,
          paymentAmount,
          docName,
          { refCount: references.length }
        );
        
        const docUrl = `${activeProfile.url}/app/payment-entry/${docName}`;
        logToConsole(`👉 Review Payment: ${docUrl}`, 'info');
        setPaymentRefNo('');
        const today = new Date().toISOString().split('T')[0];
        setPaymentRefDate(today);
        setPaymentType('Pay');
        alert(`Success! Created and Submitted Payment Entry ${docName}.`);
        loadUnpaidInvoices(paySelectedSupplier.name);
      }
    } catch (error) {
      logToConsole(`✗ Payment Sync Failed: ${error.message}`, 'err');
      alert(`Sync Failed: ${error.message}`);
    }
  }

  const copyRawText = () => {
    navigator.clipboard.writeText(rawText);
    alert("Copied to clipboard!");
  };

  const downloadWordDoc = () => {
    const htmlContent = rawText.replace(/\n/g, '<br/>');
    const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><title>Export Invoice Text</title><style>body { font-family: 'Courier New', Courier, monospace; font-size: 11pt; line-height: 1.5; }</style></head><body>";
    const footer = "</body></html>";
    const sourceHTML = header + htmlContent + footer;
    
    const blob = new Blob(['\ufeff' + sourceHTML], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    document.body.appendChild(a);
    a.href = url;
    a.download = 'invoice_text.doc';
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Quarterly Reconciliation Handler
  async function handleReconFileUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    setReconLoading(true);
    setReconFileStatus(`Analyzing document "${file.name}" with Gemini Vision OCR...`);
    setReconFileColor("var(--color-secondary)");

    try {
      // Convert file to Base64
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const imageSrc = event.target.result;
          const base64Parts = imageSrc.split(',');
          const mimeType = base64Parts[0].match(/:(.*?);/)?.[1] || 'image/png';
          const base64Data = base64Parts[1];

          if (!parserSettings.geminiKey) {
            throw new Error("Gemini API Key is missing. Please configure it in the Parser AI tab.");
          }
          const cleanKey = parserSettings.geminiKey.trim();

          const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${cleanKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{
                parts: [
                  { text: `Extract all invoice numbers, bill numbers, and sales order numbers from this payment list document. Return ONLY a valid JSON array of strings matching this schema: ["INV-001", "ST251103022", "PO-1029"]` },
                  { inlineData: { mimeType, data: base64Data } }
                ]
              }],
              generationConfig: { responseMimeType: 'application/json' }
            })
          });

          if (!response.ok) {
            const errData = await response.json().catch(() => null);
            throw new Error(errData?.error?.message || `HTTP ${response.status}`);
          }

          const resData = await response.json();
          const textResponse = resData.candidates[0].content.parts[0].text.trim();
          const extractedInvs = JSON.parse(textResponse);

          if (!Array.isArray(extractedInvs) || extractedInvs.length === 0) {
            setReconFileStatus("No invoice numbers were detected in the uploaded payment list.");
            setReconFileColor("var(--status-warn)");
            setReconLoading(false);
            return;
          }

          setReconInvoices(extractedInvs);
          setReconFileStatus(`Found ${extractedInvs.length} invoice code(s). Fetching rebate records from Neon DB...`);

          // Query Neon DB for matching rebate records
          const recRes = await fetch(`/api/rebates/records?company_profile=${encodeURIComponent(activeProfileName)}&invoices=${encodeURIComponent(extractedInvs.join(','))}`);
          const recData = await recRes.json();

          if (recData.success) {
            setReconRecords(recData.data || []);
            setReconFileStatus(`✓ Successfully reconciled ${recData.data.length} rebate item(s) across ${extractedInvs.length} invoice(s)!`);
            setReconFileColor("var(--status-success)");
          } else {
            throw new Error(recData.error || "Failed to query rebate records.");
          }
        } catch (err) {
          setReconFileStatus(`Error processing payment list: ${err.message}`);
          setReconFileColor("var(--status-error)");
        } finally {
          setReconLoading(false);
        }
      };
      reader.readAsDataURL(file);
    } catch (err) {
      setReconFileStatus(`File read error: ${err.message}`);
      setReconFileColor("var(--status-error)");
      setReconLoading(false);
    }
  }

  // Export Quarterly Reconciliation to Excel
  function exportReconToExcel() {
    if (reconRecords.length === 0) {
      alert("No rebate records available to export.");
      return;
    }

    const exportRows = reconRecords.map(r => ({
      "Date": r.date ? new Date(r.date).toISOString().split('T')[0] : '',
      "Invoice Number": r.invoice_no,
      "Item Model Code": r.item_model_code,
      "Quantity": parseFloat(r.quantity),
      "Original Rate ($)": parseFloat(r.rate),
      "Real Rate ($)": parseFloat(r.real_rate),
      "Rebate %": (parseFloat(r.rebate_percentage) * 100).toFixed(1) + "%",
      "Rebate Amount ($)": parseFloat(r.rebate_amount),
      "Company Profile": r.company_profile
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Rebate Reconciliation");
    XLSX.writeFile(workbook, `Quarterly_Rebate_Reconciliation_${activeProfileName || 'Export'}.xlsx`);
  }

  async function deleteRebateRecord(id) {
    if (!confirm("Are you sure you want to permanently delete this rebate record?")) return;
    try {
      const res = await fetch(`/api/rebates/records?id=${id}`, { method: 'DELETE' });
      const data = await res.json();
      if (data.success) {
        setReconRecords(prev => prev.filter(r => r.id !== id));
        alert("Record deleted successfully.");
      } else {
        alert("Failed to delete record: " + data.error);
      }
    } catch (err) {
      alert("Error deleting record: " + err.message);
    }
  }

  async function loadAllRebateRecords() {
    setReconLoading(true);
    setReconFileStatus("Loading all database records...");
    setReconFileColor("var(--text-muted)");
    try {
      const recRes = await fetch(`/api/rebates/records?company_profile=${encodeURIComponent(activeProfileName)}`);
      const recData = await recRes.json();
      if (recData.success) {
        setReconRecords(recData.data || []);
        setReconFileStatus(`Loaded all ${recData.data.length} rebate records from database.`);
        setReconFileColor("var(--status-success)");
      } else {
        throw new Error(recData.error || "Failed to load records.");
      }
    } catch (err) {
      setReconFileStatus(`Error loading records: ${err.message}`);
      setReconFileColor("var(--status-error)");
    } finally {
      setReconLoading(false);
    }
  }

  // Multi-tab Excel Exception List Upload Handler
  async function handleExceptionListUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    setExceptionUploadStatus(`Parsing all tabs in Excel file "${file.name}"...`);
    setExceptionUploadColor("var(--color-secondary)");

    try {
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array' });

          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error("The uploaded Excel file contains no worksheets.");
          }

          const extractedModelsSet = new Set();
          let totalSheetsProcessed = 0;

          // Process ALL worksheets/tabs in the Excel workbook
          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) return;

            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            if (!jsonData || jsonData.length === 0) return;

            totalSheetsProcessed++;

            // Find column index for MODEL in this worksheet
            let modelColIndex = -1;
            let startRowIndex = 1;
            const headerRow = jsonData[0] || [];

            headerRow.forEach((cell, idx) => {
              const str = String(cell || '').toUpperCase().trim();
              if (
                str === 'MODEL' ||
                str === 'MODEL CODE' ||
                str === 'ITEM CODE' ||
                str === 'ITEM' ||
                str === 'MODEL_NO' ||
                str === 'MODELS'
              ) {
                modelColIndex = idx;
              }
            });

            if (modelColIndex === -1) {
              modelColIndex = 0; // Default to Column A
              const firstCellStr = String(headerRow[0] || '').toUpperCase().trim();
              if (!['MODEL', 'MODEL CODE', 'ITEM CODE', 'ITEM', 'MODEL_NO', 'MODELS'].includes(firstCellStr)) {
                startRowIndex = 0;
              }
            }

            for (let r = startRowIndex; r < jsonData.length; r++) {
              const row = jsonData[r];
              if (row && row[modelColIndex] !== undefined && row[modelColIndex] !== null) {
                const rawVal = String(row[modelColIndex]).trim();
                const upperVal = rawVal.toUpperCase();

                // STRICT PREFIX FILTER: Only extract model codes starting with DS, iDS, CS, RG, or VS
                const isTargetPrefix =
                  upperVal.startsWith('DS') ||
                  upperVal.startsWith('IDS') ||
                  upperVal.startsWith('CS') ||
                  upperVal.startsWith('RG') ||
                  upperVal.startsWith('VS');

                // Skip header re-definitions, non-matching prefixes, empty cells, and noise text
                if (
                  rawVal.length > 0 &&
                  isTargetPrefix &&
                  !['MODEL', 'MODEL CODE', 'ITEM CODE', 'ITEM', 'DESCRIPTION', 'PICTURE', 'PRICE', 'PAGE'].includes(upperVal)
                ) {
                  extractedModelsSet.add(rawVal);
                  
                  // Also extract clean prefix code if model contains parenthetical text e.g. "DS-TMG4B0-RA(4m)" -> "DS-TMG4B0-RA"
                  const cleanCodeMatch = rawVal.match(/^([A-Z0-9\-_]+)/i);
                  if (cleanCodeMatch && cleanCodeMatch[1] && cleanCodeMatch[1].length >= 3) {
                    const cleanUpper = cleanCodeMatch[1].toUpperCase();
                    if (
                      cleanUpper.startsWith('DS') ||
                      cleanUpper.startsWith('IDS') ||
                      cleanUpper.startsWith('CS') ||
                      cleanUpper.startsWith('RG') ||
                      cleanUpper.startsWith('VS')
                    ) {
                      extractedModelsSet.add(cleanCodeMatch[1].trim());
                    }
                  }
                }
              }
            }
          });

          const extractedModels = Array.from(extractedModelsSet);

          if (extractedModels.length === 0) {
            throw new Error(`Processed ${totalSheetsProcessed} tab(s), but no valid model codes starting with DS, iDS, CS, RG, or VS were found.`);
          }

          setExceptionUploadStatus(`Uploading ${extractedModels.length} exception model code(s) from ${totalSheetsProcessed} tab(s) to Neon DB...`);

          const res = await fetch('/api/rebates/exceptions', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ items: extractedModels })
          });

          const resData = await res.json();
          if (resData.success) {
            const msg = `Successfully updated exception masterlist with ${extractedModels.length} items from ${totalSheetsProcessed} tab(s).`;
            setExceptionUploadStatus(`✓ ${msg}`);
            setExceptionUploadColor("var(--status-success)");
            alert(msg);
            fetchRebateExceptionList();
          } else {
            throw new Error(resData.error || "Failed to update exception list.");
          }
        } catch (err) {
          setExceptionUploadStatus(`Upload failed: ${err.message}`);
          setExceptionUploadColor("var(--status-error)");
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      setExceptionUploadStatus(`File read error: ${err.message}`);
      setExceptionUploadColor("var(--status-error)");
    }
  }

  function formatDate(dStr) {
    if (!dStr) return '';
    const clean = dStr.replace(/[\/\.]/g, '-');
    const match = clean.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (match) return clean;
    const matchDMY = clean.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
    if (matchDMY) return `${matchDMY[3]}-${matchDMY[2].padStart(2, '0')}-${matchDMY[1].padStart(2, '0')}`;
    return clean;
  }

  return (
    <div className="app-container">
      {/* Header */}
      <header className="app-header">
        <div className="logo-area">
          <span className="logo-icon">⚡</span>
          <div className="logo-text">
            <h1>Invoice Sync Automation Portal</h1>
            <span className="version-tag">ERPNext v16 & Neon DB Standalone Web System</span>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '15px', alignItems: 'center' }}>
          {activeProfile && (
            <span style={{ fontSize: '12px', fontWeight: '500', color: 'var(--text-muted)' }}>
              Active Site: <strong style={{ color: 'var(--text-main)' }}>{activeProfileName}</strong>
            </span>
          )}
          {connectionStatus === 'connected' && <span className="badge badge-connected">● Connected</span>}
          {connectionStatus === 'testing' && <span className="badge badge-testing">⚡ Testing</span>}
          {connectionStatus === 'disconnected' && <span className="badge badge-disconnected">○ Offline</span>}
        </div>
      </header>

      {/* Tabs */}
      <nav className="tab-navigation">
        <button className={`tab-btn ${activeTab === 'dashboard' ? 'active' : ''}`} onClick={() => setActiveTab('dashboard')}>📊 Dashboard & Logs</button>
        <button className={`tab-btn ${activeTab === 'parser' ? 'active' : ''}`} onClick={() => setActiveTab('parser')}>📄 Invoice Parser</button>
        <button className={`tab-btn ${activeTab === 'payment' ? 'active' : ''}`} onClick={() => setActiveTab('payment')}>💳 Payment Entry</button>
        <button className={`tab-btn ${activeTab === 'reconciliation' ? 'active' : ''}`} onClick={() => setActiveTab('reconciliation')}>📈 Rebate Reconciliation</button>
        <button className={`tab-btn ${activeTab === 'rebate-settings' ? 'active' : ''}`} onClick={() => setActiveTab('rebate-settings')}>⚙️ Rebate Exceptions</button>
        <button className={`tab-btn ${activeTab === 'settings' ? 'active' : ''}`} onClick={() => setActiveTab('settings')}>🏢 ERPNext Profiles</button>
        <button className={`tab-btn ${activeTab === 'parser-settings' ? 'active' : ''}`} onClick={() => setActiveTab('parser-settings')}>🤖 Parser AI</button>
      </nav>

      {/* Panel 1: Dashboard */}
      {activeTab === 'dashboard' && (
        connectionStatus === 'connected' ? (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            <div className="analytics-grid">
              <div className="analytic-card">
                <span className="analytic-title">Invoices Synced</span>
                <span className="analytic-val">{metrics.total_count}</span>
                <span className="analytic-sub">✓ Audit Trail Saved</span>
              </div>
              <div className="analytic-card">
                <span className="analytic-title">Total Value Synced</span>
                <span className="analytic-val">${metrics.total_synced_amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                <span className="analytic-sub">Via Vercel Serverless</span>
              </div>
              <div className="analytic-card">
                <span className="analytic-title">Total Payments Issued</span>
                <span className="analytic-val">${metrics.total_paid_amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                <span className="analytic-sub">Neon DB Logged</span>
              </div>
            </div>

            <div className="form-card">
              <h2>System Activity & Logs (Neon Postgres Audit Trail)</h2>
              <div className="table-container">
                <table className="items-table">
                  <thead>
                    <tr>
                      <th>Timestamp</th>
                      <th>Action</th>
                      <th>Profile</th>
                      <th>Supplier</th>
                      <th>Amount</th>
                      <th>ERP Doc ID</th>
                    </tr>
                  </thead>
                  <tbody>
                    {auditLogs.map((log) => (
                      <tr key={log.id}>
                        <td>{new Date(log.created_at).toLocaleString()}</td>
                        <td>
                          <span style={{ 
                            color: log.action_type === 'SYNC_SUCCESS' ? 'var(--status-success)' : 'var(--color-secondary)',
                            fontWeight: '600'
                          }}>
                            {log.action_type}
                          </span>
                        </td>
                        <td>{log.profile_name}</td>
                        <td>{log.supplier || 'N/A'}</td>
                        <td>${parseFloat(log.amount).toFixed(2)}</td>
                        <td>{log.doc_name || 'N/A'}</td>
                      </tr>
                    ))}
                    {auditLogs.length === 0 && (
                      <tr>
                        <td colSpan="6" style={{ textAlign: 'center', color: 'var(--text-muted)' }}>No audit history found.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : (
          <div className="form-card" style={{ 
            textAlign: 'center', 
            padding: '60px 20px', 
            background: 'rgba(30, 41, 59, 0.4)',
            backdropFilter: 'blur(8px)',
            border: '1px solid rgba(255, 255, 255, 0.05)',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            gap: '15px',
            minHeight: '320px'
          }}>
            <div style={{ fontSize: '48px', marginBottom: '10px' }}>🔒</div>
            <h2 style={{ margin: 0, color: 'var(--color-primary)' }}>Dashboard & Logs Restricted</h2>
            <p style={{ color: 'var(--text-muted)', maxWidth: '460px', fontSize: '13.5px', lineHeight: '1.6', margin: 0 }}>
              Audit trails and sync metrics are restricted. Please log in and connect to your ERPNext profile under the <strong>ERPNext Profiles</strong> tab to unlock the dashboard.
            </p>
            <button 
              className="btn btn-primary" 
              style={{ marginTop: '10px', padding: '10px 28px' }}
              onClick={() => setActiveTab('settings')}
            >
              Go to ERPNext Profiles
            </button>
          </div>
        )
      )}

      {/* Panel 2: Invoice Parser */}
      {activeTab === 'parser' && (
        <div className="grid-two-columns">
          {/* Left: Input, Logs, Extracted Text */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            <div className="form-card">
              <h2>Upload Invoice Soft Copy</h2>
              <div 
                className="drop-zone"
                onClick={triggerFileSelect}
                onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add('dragover'); }}
                onDragLeave={(e) => e.currentTarget.classList.remove('dragover')}
                onDrop={(e) => { e.preventDefault(); e.currentTarget.classList.remove('dragover'); handleFileDropChange(e); }}
              >
                <div className="drop-zone-icon">📥</div>
                <p className="drop-zone-text">Drag & drop PDF / image invoice here</p>
                <span className="drop-zone-or">or</span>
                <button type="button" className="btn btn-secondary">Choose File</button>
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  onChange={handleFileDropChange} 
                  accept="image/*,application/pdf" 
                  style={{ display: 'none' }} 
                />
              </div>

              {progress.show && (
                <div className="status-card">
                  <div className="spinner-container">
                    <div className="double-bounce1"></div>
                    <div className="double-bounce2"></div>
                  </div>
                  <div className="status-details">
                    <h3>{progress.title}</h3>
                    <div className="progress-bar-bg">
                      <div className="progress-bar-fill" style={{ width: `${progress.percent}%` }}></div>
                    </div>
                    <p>{progress.details}</p>
                  </div>
                </div>
              )}
            </div>

            {/* Extracted Text */}
            {showTextSection && (
              <div className="form-card" style={{ padding: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                  <span style={{ fontSize: '12px', fontWeight: '700', color: 'var(--color-secondary)' }}>📄 Raw Extracted Text</span>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button type="button" className="btn btn-secondary btn-small" onClick={copyRawText}>Copy</button>
                    <button type="button" className="btn btn-secondary btn-small" onClick={downloadWordDoc}>Word Doc</button>
                  </div>
                </div>
                <textarea 
                  value={rawText} 
                  onChange={(e) => setRawText(e.target.value)}
                  style={{ height: '120px', fontFamily: 'monospace', fontSize: '11px', backgroundColor: 'rgba(0,0,0,0.3)' }}
                />
              </div>
            )}

            {/* Logs console */}
            <div className="console-card">
              <div className="console-header">
                <span>Execution Logs</span>
                <button className="btn btn-secondary btn-small" style={{ padding: '2px 8px', fontSize: '10px' }} onClick={() => setLogs([])}>Clear</button>
              </div>
              <div className="console-body">
                {logs.map((l, index) => (
                  <div key={index} className={`log-line log-${l.type}`}>
                    [{l.timestamp}] {l.text}
                  </div>
                ))}
                {logs.length === 0 && <div style={{ color: '#4b5563' }}>Portal ready. Upload an invoice to begin.</div>}
                <div ref={logsEndRef} />
              </div>
            </div>
          </div>

          {/* Right: Verification Form */}
          <div className="form-card">
            <h2>Verify Parsed Invoice Data</h2>
            <form onSubmit={(e) => { e.preventDefault(); syncInvoiceToERPNext(); }}>
              <div className="form-group" style={{ position: 'relative' }}>
                <label>Supplier / Party <span className="required">*</span></label>
                <div className="autocomplete-container">
                  <input 
                    type="text" 
                    placeholder="Search or select supplier..."
                    value={supplierQuery}
                    onChange={(e) => {
                      setSupplierQuery(e.target.value);
                      setShowSupplierDropdown(true);
                      const q = e.target.value.toLowerCase();
                      const filtered = suppliersList.filter(s => (s.supplier_name || s.name || "").toLowerCase().includes(q) || s.name.toLowerCase().includes(q));
                      setSupplierSuggestions(filtered);
                    }}
                    onFocus={() => {
                      if (!selectedSupplierObject) {
                        setShowSupplierDropdown(true);
                        setSupplierSuggestions(suppliersList);
                      }
                    }}
                    required
                  />
                  {showSupplierDropdown && (
                    <div className="autocomplete-dropdown">
                      {suppliersList.length === 0 ? (
                        <div className="autocomplete-item" style={{ color: 'var(--status-warn)', cursor: 'default' }}>
                          ⚠️ No suppliers loaded. Ensure ERPNext profile is active and status is Connected.
                        </div>
                      ) : supplierSuggestions.length === 0 ? (
                        <div className="autocomplete-item" style={{ color: 'var(--text-muted)', cursor: 'default' }}>
                          No matches found.
                        </div>
                      ) : (
                        supplierSuggestions.map(s => (
                          <div 
                            key={s.name}
                            className="autocomplete-item"
                            onClick={() => {
                              setSelectedSupplierObject(s);
                              setSupplierQuery(s.supplier_name || s.name || "");
                              setShowSupplierDropdown(false);
                              setInvoiceForm(prev => ({ ...prev, supplier: s.name }));
                            }}
                          >
                            {s.supplier_name || s.name} <span style={{ fontSize: '10px', opacity: 0.6 }}>({s.name})</span>
                          </div>
                        ))
                      )}
                    </div>
                  )}
                </div>
                <small style={{ 
                  fontSize: '11px', 
                  color: selectedSupplierObject ? 'var(--status-success)' : 'var(--status-warn)',
                  fontWeight: '600'
                }}>
                  {selectedSupplierObject ? `✓ Matched with ERPNext: ${selectedSupplierObject.name}` : '⚠ Not matched with ERPNext supplier.'}
                </small>
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Invoice Number (Bill No) <span className="required">*</span></label>
                  <input 
                    type="text" 
                    value={invoiceForm.invoice_number} 
                    onChange={(e) => setInvoiceForm({ ...invoiceForm, invoice_number: e.target.value })} 
                    required 
                  />
                </div>
                <div className="form-group">
                  <label>Invoice Date <span className="required">*</span></label>
                  <input 
                    type="date" 
                    value={invoiceForm.date} 
                    onChange={(e) => setInvoiceForm({ ...invoiceForm, date: e.target.value })} 
                    required 
                  />
                </div>
              </div>

              <div style={{ marginTop: '20px', marginBottom: '15px' }}>
                {/* Chat Prices Input Field */}
                <div className="form-group" style={{ marginBottom: '16px' }}>
                  <label>Chat Prices (comma-separated, e.g. 62,30)</label>
                  <input 
                    type="text" 
                    placeholder="Enter agreed prices in sequence: 62, 30, 45"
                    value={chatPricesInput}
                    onChange={(e) => {
                      const val = e.target.value;
                      setChatPricesInput(val);
                      const updatedItems = computeLineItemRealRates(val, invoiceForm.items);
                      setInvoiceForm({ ...invoiceForm, items: updatedItems });
                    }}
                  />
                  <small style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                    Sequentially maps prices to physical items. Unmapped items default to original invoice rate.
                  </small>
                </div>

                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                  <label style={{ fontSize: '11px' }}>Line Items Table</label>
                  <button 
                    type="button" 
                    className="btn btn-secondary btn-small"
                    onClick={() => {
                      const newItems = [...invoiceForm.items, { description: '', qty: 1, rate: 0, amount: 0, real_rate: 0, item_code: '' }];
                      const mapped = computeLineItemRealRates(chatPricesInput, newItems);
                      setInvoiceForm({ ...invoiceForm, items: mapped });
                    }}
                  >
                    + Add Row
                  </button>
                </div>
                
                <div className="table-container" style={{ overflowX: 'auto', borderRadius: 'var(--radius-sm)' }}>
                  <table className="items-table" style={{ minWidth: '720px', width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr>
                        <th style={{ width: '32%', minWidth: '190px' }}>Item Code/Name</th>
                        <th style={{ width: '9%', minWidth: '60px', textAlign: 'center' }}>Qty</th>
                        <th style={{ width: '14%', minWidth: '80px', textAlign: 'right' }}>Rate</th>
                        <th style={{ width: '14%', minWidth: '80px', textAlign: 'right' }}>Real Rate</th>
                        <th style={{ width: '14%', minWidth: '95px', textAlign: 'center' }}>Rebate %</th>
                        <th style={{ width: '12%', minWidth: '75px', textAlign: 'right' }}>Rebate ($)</th>
                        <th style={{ width: '5%', minWidth: '35px' }}></th>
                      </tr>
                    </thead>
                    <tbody>
                      {invoiceForm.items.map((item, index) => {
                        const reb = getItemRebateInfo(item);
                        return (
                          <tr key={index}>
                            <td style={{ verticalAlign: 'top', padding: '8px 6px' }}>
                              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                                  <span style={{ fontSize: '10px', color: 'var(--text-muted)', width: '68px', flexShrink: 0 }}>Invoice Desc:</span>
                                  <input 
                                    type="text" 
                                    value={item.description} 
                                    onChange={(e) => {
                                      const newItems = [...invoiceForm.items];
                                      newItems[index].description = e.target.value;
                                      const matchResult = findMatchingItemCode(e.target.value);
                                      newItems[index].item_code = matchResult.item_code;
                                      newItems[index].match_confidence = matchResult.match_confidence;
                                      setInvoiceForm({ ...invoiceForm, items: newItems });
                                    }}
                                    style={{ padding: '4px 8px', fontSize: '12px', flex: 1, minWidth: '110px' }}
                                    placeholder="Description on invoice..."
                                  />
                                </div>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                                  <span style={{ fontSize: '10px', color: 'var(--color-secondary)', width: '68px', flexShrink: 0, fontWeight: '600' }}>ERP Item:</span>
                                  <input 
                                    type="text" 
                                    value={item.item_code || ''} 
                                    onChange={(e) => {
                                      const newItems = [...invoiceForm.items];
                                      newItems[index].item_code = e.target.value;
                                      newItems[index].match_confidence = 'high'; // Assume manual selection is high confidence
                                      setInvoiceForm({ ...invoiceForm, items: newItems });
                                    }}
                                    style={{ padding: '4px 8px', fontSize: '12px', flex: 1, minWidth: '110px', border: '1px solid rgba(56, 189, 248, 0.3)' }}
                                    placeholder="Type or select item code..."
                                    list="erp-items-datalist"
                                  />
                                </div>
                                {(() => {
                                  const isDiscount = parseFloat(item.rate || 0) < 0 || (item.description || '').toLowerCase().includes('discount');
                                  if (isDiscount) return null;
                                  
                                  const erpItemMatch = itemsList.find(it => (it.item_code || it.name || '').toLowerCase() === (item.item_code || '').toLowerCase());
                                  
                                  if (!erpItemMatch && item.item_code && item.item_code !== 'Default Item') {
                                    return (
                                      <div style={{ fontSize: '10px', color: '#fbbf24', display: 'flex', alignItems: 'center', gap: '4px', marginTop: '2px', fontWeight: '500', lineHeight: '1.2' }}>
                                        ⚠️ New Item (Not in ERPNext)
                                      </div>
                                    );
                                  } else if (erpItemMatch && item.match_confidence === 'low') {
                                    return (
                                      <div style={{ fontSize: '10px', color: '#fbbf24', display: 'flex', alignItems: 'center', gap: '4px', marginTop: '2px', fontWeight: '500', lineHeight: '1.2' }}>
                                        ⚠️ Low Confidence Match: Please Verify
                                      </div>
                                    );
                                  }

                                  const lastPurchaseRate = erpItemMatch?.last_purchase_rate || 0;
                                  const effectiveRealRate = item.real_rate !== undefined ? parseFloat(item.real_rate) : parseFloat(item.rate || 0);
                                  
                                  if (lastPurchaseRate > 0 && effectiveRealRate > lastPurchaseRate) {
                                    return (
                                      <div style={{ 
                                        fontSize: '10px', 
                                        color: 'var(--status-error)', 
                                        display: 'flex', 
                                        alignItems: 'center', 
                                        gap: '4px', 
                                        marginTop: '2px',
                                        fontWeight: '500',
                                        lineHeight: '1.2'
                                      }}>
                                        ⚠️ Exceeds last price (${lastPurchaseRate.toFixed(2)})
                                      </div>
                                    );
                                  }
                                  return null;
                                })()}
                                {profiles.find(p => p.profile_name === activeProfileName)?.enable_sn_tracking && (
                                  <div style={{ display: 'flex', alignItems: 'flex-start', gap: '4px', marginTop: '2px' }}>
                                    <span style={{ fontSize: '10px', color: 'var(--text-muted)', width: '68px', flexShrink: 0, marginTop: '4px' }}>Serial Nos:</span>
                                    <textarea 
                                      value={item.serial_nos || ''} 
                                      onChange={(e) => {
                                        const newItems = [...invoiceForm.items];
                                        newItems[index].serial_nos = e.target.value;
                                        setInvoiceForm({ ...invoiceForm, items: newItems });
                                      }}
                                      style={{ padding: '4px 8px', fontSize: '11px', flex: 1, minWidth: '110px', minHeight: '30px', resize: 'vertical' }}
                                      placeholder="Paste SNs..."
                                    />
                                  </div>
                                )}
                              </div>
                            </td>
                            <td style={{ verticalAlign: 'top', padding: '8px 4px' }}>
                              <input 
                                type="number" 
                                value={item.qty} 
                                onChange={(e) => {
                                  const newItems = [...invoiceForm.items];
                                  newItems[index].qty = parseFloat(e.target.value || 0);
                                  newItems[index].amount = newItems[index].qty * newItems[index].rate;
                                  setInvoiceForm({ ...invoiceForm, items: newItems });
                                }}
                                style={{ padding: '6px 4px', fontSize: '12.5px', textAlign: 'center', width: '100%', minWidth: '55px', boxSizing: 'border-box' }}
                              />
                            </td>
                            <td style={{ verticalAlign: 'top', padding: '8px 4px' }}>
                              <input 
                                type="number" 
                                step="0.01"
                                value={item.rate} 
                                onChange={(e) => {
                                  const newItems = [...invoiceForm.items];
                                  newItems[index].rate = parseFloat(e.target.value || 0);
                                  newItems[index].amount = newItems[index].qty * newItems[index].rate;
                                  const remapped = computeLineItemRealRates(chatPricesInput, newItems);
                                  setInvoiceForm({ ...invoiceForm, items: remapped });
                                }}
                                style={{ padding: '6px 6px', fontSize: '12.5px', textAlign: 'right', width: '100%', minWidth: '70px', boxSizing: 'border-box' }}
                              />
                            </td>
                            <td style={{ verticalAlign: 'top', padding: '8px 4px' }}>
                              <input 
                                type="number" 
                                step="0.01"
                                value={item.real_rate !== undefined ? item.real_rate : item.rate} 
                                onChange={(e) => {
                                  const newItems = [...invoiceForm.items];
                                  newItems[index].real_rate = parseFloat(e.target.value || 0);
                                  setInvoiceForm({ ...invoiceForm, items: newItems });
                                }}
                                style={{ padding: '6px 6px', fontSize: '12.5px', textAlign: 'right', width: '100%', minWidth: '70px', boxSizing: 'border-box', borderColor: 'rgba(56, 189, 248, 0.4)' }}
                              />
                            </td>
                            <td style={{ verticalAlign: 'top', padding: '8px 4px', textAlign: 'center' }}>
                              <span 
                                className="badge" 
                                style={{ 
                                  fontSize: '10.5px',
                                  padding: '4px 6px',
                                  display: 'inline-block',
                                  minWidth: '85px',
                                  textAlign: 'center',
                                  backgroundColor: reb.pct > 0 ? (reb.pct === 0.08 ? 'rgba(34, 197, 94, 0.15)' : 'rgba(56, 189, 248, 0.15)') : 'rgba(255, 255, 255, 0.05)', 
                                  color: reb.pct > 0 ? (reb.pct === 0.08 ? 'var(--status-success)' : 'var(--color-secondary)') : 'var(--text-muted)' 
                                }}
                              >
                                {(reb.pct * 100).toFixed(1)}% <br/>
                                <span style={{ fontSize: '9px', opacity: 0.85 }}>{reb.label}</span>
                              </span>
                            </td>
                            <td style={{ verticalAlign: 'top', padding: '10px 4px', textAlign: 'right' }}>
                              <strong style={{ color: reb.amt > 0 ? 'var(--status-success)' : 'var(--text-muted)', fontSize: '12.5px' }}>
                                ${reb.amt.toFixed(2)}
                              </strong>
                            </td>
                            <td style={{ verticalAlign: 'top', padding: '8px 4px', textAlign: 'center' }}>
                              <button 
                                type="button" 
                                className="btn-delete-row" 
                                onClick={() => {
                                  const newItems = invoiceForm.items.filter((_, i) => i !== index);
                                  const remapped = computeLineItemRealRates(chatPricesInput, newItems);
                                  setInvoiceForm({ ...invoiceForm, items: remapped });
                                }}
                              >
                                🗑️
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

              {(() => {
                const expectedDisc = computeExpectedDiscount(invoiceForm.items);
                const actualDisc = parseFloat(invoiceForm.discount_amount || 0);
                const isDiscMismatch = Math.abs(expectedDisc - actualDisc) > 0.02;
                const totalRebateSum = invoiceForm.items.reduce((sum, item) => sum + getItemRebateInfo(item).amt, 0);

                return (
                  <div className="totals-section">
                    <div className="total-row">
                      <span>Calculated Items Total:</span>
                      <span className="total-value">
                        ${invoiceForm.items.reduce((sum, item) => sum + (item.qty * item.rate), 0).toFixed(2)}
                      </span>
                    </div>
                    <div className="total-row" style={{ color: 'var(--status-success)' }}>
                      <span>Total Estimated Rebates:</span>
                      <span className="total-value" style={{ color: 'var(--status-success)' }}>
                        ${totalRebateSum.toFixed(2)}
                      </span>
                    </div>
                    <div className="form-row">
                      <div className="form-group">
                        <label>Discount Amount (Global)</label>
                        <input 
                          type="number" 
                          step="0.01"
                          value={invoiceForm.discount_amount} 
                          onChange={(e) => setInvoiceForm({ ...invoiceForm, discount_amount: parseFloat(e.target.value || 0) })} 
                          style={isDiscMismatch ? {
                            borderColor: '#ef4444',
                            backgroundColor: 'rgba(239, 68, 68, 0.15)',
                            color: '#fca5a5',
                            fontWeight: '700'
                          } : {}}
                        />
                        {isDiscMismatch && (
                          <small style={{ color: '#fca5a5', fontWeight: '600', display: 'block', marginTop: '4px' }}>
                            ⚠️ Discount Mismatch! Expected ${expectedDisc.toFixed(2)} based on Real Rates, but Discount Amount is ${actualDisc.toFixed(2)}.
                          </small>
                        )}
                      </div>
                      <div className="form-group">
                        <label>Grand Total Target</label>
                        <input 
                          type="number" 
                          step="0.01"
                          value={invoiceForm.grand_total} 
                          onChange={(e) => setInvoiceForm({ ...invoiceForm, grand_total: parseFloat(e.target.value || 0) })} 
                        />
                      </div>
                    </div>
                  </div>
                );
              })()}

              {activeProfile?.sync_doctype === 'Purchase Order' && (
                <div style={{ marginTop: '16px', padding: '12px', border: '1px solid rgba(56, 189, 248, 0.2)', borderRadius: 'var(--radius-sm)', backgroundColor: 'rgba(56, 189, 248, 0.02)' }}>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '13px', fontWeight: '600' }}>
                    <input 
                      type="checkbox" 
                      checked={autoSubmitFlow} 
                      onChange={(e) => setAutoSubmitFlow(e.target.checked)} 
                      style={{ width: '16px', height: '16px', margin: 0 }}
                    />
                    <span>⚡ Auto-Submit PO and Auto-Generate Receipt & Invoice</span>
                  </label>
                  <small style={{ display: 'block', fontSize: '11px', color: 'var(--text-muted)', marginTop: '4px', marginLeft: '24px', lineHeight: '1.4' }}>
                    When checked, the system will submit the PO, verify inventory receipt, and post the purchase invoice with the supplier invoice number automatically in ERPNext.
                  </small>
                </div>
              )}
              
              <div className="form-group" style={{ marginBottom: '24px' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', margin: 0 }}>
                  <input 
                    type="checkbox" 
                    checked={taxInclusive}
                    onChange={(e) => setTaxInclusive(e.target.checked)}
                    style={{ width: '16px', height: '16px' }}
                  />
                  <span>Tax Conclusive (Prices include GST)</span>
                </label>
                <small style={{ display: 'block', fontSize: '11px', color: 'var(--text-muted)', marginTop: '4px', marginLeft: '24px', lineHeight: '1.4' }}>
                  When checked, the system will divide the parsed amounts by the tax rate to extract the base net total before applying taxes in ERPNext.
                </small>
              </div>


              <button 
                type="submit" 
                className="btn btn-primary btn-full" 
                style={{ marginTop: '16px' }}
                disabled={connectionStatus !== 'connected'}
              >
                ⚡ Sync Draft {activeProfile?.sync_doctype || 'Purchase Order'} to ERPNext
              </button>
            </form>
          </div>
        </div>
      )}

      {/* Panel 3: Payment Entry */}
      {activeTab === 'payment' && (
        <div className="form-card" style={{ maxWidth: '800px', margin: '0 auto' }}>
          <h2>Create Supplier Payment Entry</h2>
          <form onSubmit={(e) => { e.preventDefault(); submitPaymentEntry(); }}>
            <div className="form-group" style={{ position: 'relative' }}>
              <label>Select Supplier <span className="required">*</span></label>
              <div className="autocomplete-container">
                <input 
                  type="text" 
                  placeholder="Search or select supplier..."
                  value={paySupplierQuery}
                  onChange={(e) => {
                    setPaySupplierQuery(e.target.value);
                    setShowPaySupplierDropdown(true);
                    const q = e.target.value.toLowerCase();
                    const filtered = suppliersList.filter(s => (s.supplier_name || s.name || "").toLowerCase().includes(q) || s.name.toLowerCase().includes(q));
                    setPaySupplierSuggestions(filtered);
                  }}
                  onFocus={() => {
                    setShowPaySupplierDropdown(true);
                    setPaySupplierSuggestions(suppliersList);
                  }}
                />
                {showPaySupplierDropdown && (
                  <div className="autocomplete-dropdown">
                    {suppliersList.length === 0 ? (
                      <div className="autocomplete-item" style={{ color: 'var(--status-warn)', cursor: 'default' }}>
                        ⚠️ No suppliers loaded. Ensure ERPNext profile is active and status is Connected.
                      </div>
                    ) : paySupplierSuggestions.length === 0 ? (
                      <div className="autocomplete-item" style={{ color: 'var(--text-muted)', cursor: 'default' }}>
                        No matches found.
                      </div>
                    ) : (
                      paySupplierSuggestions.map(s => (
                        <div 
                          key={s.name}
                          className="autocomplete-item"
                          onClick={() => {
                            setPaySelectedSupplier(s);
                            setPaySupplierQuery(s.supplier_name || s.name || "");
                            setShowPaySupplierDropdown(false);
                            loadUnpaidInvoices(s.name);
                          }}
                        >
                          {s.supplier_name || s.name} <span style={{ fontSize: '10px', opacity: 0.6 }}>({s.name})</span>
                        </div>
                      ))
                    )}
                  </div>
                )}
              </div>
              <small style={{ color: paymentListColor, fontWeight: '600', fontSize: '11px' }}>
                {paymentListStatus}
              </small>
            </div>

            {unpaidInvoices.length > 0 && (() => {
              const selectedCount = Object.values(checkedInvoiceIds).filter(Boolean).length;
              return (
                <div className="form-group" style={{ marginTop: '16px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                    <label style={{ margin: 0 }}>Select Unpaid Invoices to Settle</label>
                    <span style={{ 
                      fontSize: '12px', 
                      fontWeight: '700', 
                      color: selectedCount > 0 ? 'var(--status-success)' : 'var(--text-muted)',
                      backgroundColor: selectedCount > 0 ? 'rgba(16, 185, 129, 0.12)' : 'rgba(255, 255, 255, 0.05)',
                      padding: '3px 10px',
                      borderRadius: '12px',
                      border: selectedCount > 0 ? '1px solid rgba(16, 185, 129, 0.3)' : '1px solid var(--border-color)'
                    }}>
                      ✓ Selected: {selectedCount} of {unpaidInvoices.length} invoice{unpaidInvoices.length === 1 ? '' : 's'}
                    </span>
                  </div>
                  <input 
                    type="text" 
                    placeholder="🔍 Filter invoice ID or supplier No..."
                    value={payInvoiceSearch}
                    onChange={(e) => setPayInvoiceSearch(e.target.value)}
                    style={{ marginBottom: '8px', padding: '6px 12px', fontSize: '12px' }}
                  />

                  <div style={{ padding: '12px', border: '1px dashed var(--border-color)', borderRadius: 'var(--radius-sm)', textAlign: 'center', marginBottom: '8px', backgroundColor: 'rgba(255,255,255,0.01)' }}>
                    <span style={{ fontSize: '11px', fontWeight: '600', color: 'var(--text-muted)', display: 'block', marginBottom: '4px' }}>📂 Auto-check via Payment List Document/Image</span>
                    <button type="button" className="btn btn-secondary btn-small" onClick={triggerPayListSelect}>Select Payment List</button>
                    <input 
                      type="file" 
                      ref={payListInputRef} 
                      onChange={handlePayListFileChange} 
                      accept="image/*,application/pdf,text/plain" 
                      style={{ display: 'none' }} 
                    />
                  </div>

                  <div className="unpaid-invoices-container">
                    {loadingInvoices ? (
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '200px', gap: '16px' }}>
                        <div className="spinner-container">
                          <div className="double-bounce1"></div>
                          <div className="double-bounce2"></div>
                        </div>
                        <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>Loading unpaid invoices...</span>
                      </div>
                    ) : (() => {
                      const displayed = getProcessedInvoicesList();
                      const checkedAmtCounts = {};
                      const checkedBillCounts = {};
                      displayed.forEach(inv => {
                        if (checkedInvoiceIds[inv.name]) {
                          const amtKey = inv.outstanding_amount.toFixed(2);
                          checkedAmtCounts[amtKey] = (checkedAmtCounts[amtKey] || 0) + 1;
                          
                          const bill = (inv.bill_no || '').trim().toLowerCase();
                          if (bill && bill !== 'n/a') {
                            checkedBillCounts[bill] = (checkedBillCounts[bill] || 0) + 1;
                          }
                        }
                      });
                      
                      if (displayed.length === 0) {
                        return (
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100px', color: 'var(--text-muted)', fontSize: '13px' }}>
                            No unpaid invoices found.
                          </div>
                        );
                      }
                      
                      return displayed.map(inv => {
                        const isChecked = checkedInvoiceIds[inv.name] || false;
                        const isDupAmt = globalDuplicates.amountCounts[inv.outstanding_amount.toFixed(2)] > 1;
                        const isDupBill = inv.bill_no && inv.bill_no !== 'N/A' && globalDuplicates.billNoCounts[inv.bill_no.toLowerCase()] > 1;
                        const hasWarn = isDupAmt || isDupBill;
                        
                        let showHighlight = false;
                        if (isChecked) {
                          const amtKey = inv.outstanding_amount.toFixed(2);
                          const billKey = (inv.bill_no || '').trim().toLowerCase();
                          const hasCheckedAmtDup = checkedAmtCounts[amtKey] > 1;
                          const hasCheckedBillDup = billKey && billKey !== 'n/a' && checkedBillCounts[billKey] > 1;
                          showHighlight = hasCheckedAmtDup || hasCheckedBillDup;
                        } else {
                          showHighlight = hasWarn;
                        }
                        
                        return (
                          <label 
                            key={inv.name} 
                            className="invoice-checkbox-item" 
                            style={{ 
                              borderLeft: showHighlight ? '4px solid var(--status-warn)' : undefined,
                              backgroundColor: showHighlight ? 'rgba(245, 158, 11, 0.05)' : undefined 
                            }}
                          >
                            <input 
                              type="checkbox" 
                              checked={isChecked}
                              onChange={() => handleInvoiceCheckChange(inv.name, inv.outstanding_amount)}
                            />
                            <div className="invoice-checkbox-details">
                              <span className="invoice-checkbox-name">{inv.name}</span>
                              <span className="invoice-checkbox-date">
                                Supplier No: {inv.bill_no || "N/A"} | Date: {inv.posting_date}
                              </span>
                              {hasWarn && (
                                <div style={{ color: 'var(--status-warn)', fontSize: '12px', fontWeight: '600', marginTop: '6px', display: 'flex', flexWrap: 'wrap', gap: '4px', alignItems: 'center' }}>
                                  {isDupAmt && <span>⚠️ Duplicate Amount (${inv.outstanding_amount.toFixed(2)})</span>}
                                  {isDupAmt && isDupBill && <span style={{ opacity: 0.5 }}>|</span>}
                                  {isDupBill && <span>⚠️ Duplicate Bill No ({inv.bill_no.toUpperCase()})</span>}
                                </div>
                              )}
                            </div>
                            <span className="invoice-checkbox-amount">${inv.outstanding_amount.toFixed(2)}</span>
                          </label>
                        );
                      });
                    })()}
                  </div>
                </div>
              );
            })()}

            {unpaidInvoices.length > 0 && (() => {
              const selectedCount = Object.values(checkedInvoiceIds).filter(Boolean).length;
              return (
                <div style={{ marginTop: '20px', display: 'flex', flexDirection: 'column', gap: '15px' }}>
                  <div className="form-row">
                    <div className="form-group">
                      <label>Payment Type</label>
                      <select value={paymentType} onChange={(e) => setPaymentType(e.target.value)}>
                        <option value="Pay">Pay (Outflow)</option>
                        <option value="Receive">Receive (Inflow)</option>
                      </select>
                    </div>
                    <div className="form-group">
                      <label>Payment Date <span className="required">*</span></label>
                      <input type="date" value={paymentDate} onChange={(e) => setPaymentDate(e.target.value)} required />
                    </div>
                    <div className="form-group">
                      <label>Payment Method</label>
                      <select value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
                        <option value="Bank">Bank / Wire Transfer</option>
                        <option value="Cash">Cash</option>
                        <option value="Cheque">Cheque</option>
                        <option value="Credit Card">Credit Card</option>
                      </select>
                    </div>
                  </div>

                  <div className="form-row">
                    <div className="form-group">
                      <label>Cheque/Reference No. <span className="required">*</span></label>
                      <input 
                        type="text" 
                        value={paymentRefNo} 
                        onChange={(e) => setPaymentRefNo(e.target.value)} 
                        required 
                        placeholder="e.g. Chq 12345 or Inv Ref" 
                      />
                    </div>
                    <div className="form-group">
                      <label>Cheque/Reference Date <span className="required">*</span></label>
                      <input 
                        type="date" 
                        value={paymentRefDate} 
                        onChange={(e) => setPaymentRefDate(e.target.value)} 
                        required 
                      />
                    </div>
                  </div>

                  <div className="form-group">
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
                      <label style={{ margin: 0 }}>Total Payment Amount ($)</label>
                      <span style={{ fontSize: '12px', fontWeight: '700', color: selectedCount > 0 ? 'var(--status-success)' : 'var(--text-muted)' }}>
                        Selected Invoices: {selectedCount}
                      </span>
                    </div>
                    <input type="number" step="0.01" value={paymentAmount.toFixed(2)} disabled />
                    <small style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                      Auto-calculated sum for {selectedCount} selected invoice{selectedCount === 1 ? '' : 's'}.
                    </small>
                  </div>

                  <button 
                    type="submit" 
                    className="btn btn-primary btn-full"
                    disabled={connectionStatus !== 'connected' || selectedCount === 0}
                  >
                    💳 Create Payment Entry for {selectedCount} Invoice{selectedCount === 1 ? '' : 's'} (${paymentAmount.toFixed(2)})
                  </button>
                </div>
              );
            })()}
          </form>
        </div>
      )}

      {/* Panel 4: Profiles Settings */}
      {activeTab === 'settings' && (
        <div className="form-card" style={{ maxWidth: '800px', margin: '0 auto' }}>
          <h2>ERPNext Site Profiles Configuration</h2>
          <div style={{ display: 'flex', gap: '10px', marginBottom: '20px', alignItems: 'center' }}>
            <div className="form-group" style={{ flex: 1, marginBottom: 0 }}>
              <label>Select Active Website Profile</label>
              <select value={activeProfileName} onChange={(e) => setActiveProfileName(e.target.value)}>
                <option value="">-- Choose Profile --</option>
                {profiles.map(p => (
                  <option key={p.profile_name} value={p.profile_name}>{p.profile_name}</option>
                ))}
              </select>
            </div>
            <button type="button" className="btn btn-secondary" style={{ marginTop: '18px' }} onClick={handleCreateNewProfile}>＋ New</button>
            <button type="button" className="btn btn-danger" style={{ marginTop: '18px' }} onClick={deleteActiveProfile} disabled={!activeProfileName}>🗑 Delete</button>
          </div>

          {activeProfileName && (
            <form onSubmit={(e) => { e.preventDefault(); saveProfileSettings(); }}>
              <div className="form-group">
                <label>Profile Name (Label)</label>
                <input type="text" id="erp-profile_name" value={profileEditor.profile_name} disabled />
              </div>
              <div className="form-group">
                <label>ERPNext Website URL <span className="required">*</span></label>
                <input type="url" id="erp-url" value={profileEditor.url} onChange={handleEditorChange} required placeholder="https://example.erpnext.com" />
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label>Connection Routing Type</label>
                  <select id="erp-connection_type" value={profileEditor.connection_type || 'proxy'} onChange={handleEditorChange}>
                    <option value="proxy">Route via Server Proxy (Best for Cloud Sites)</option>
                    <option value="direct">Direct Browser Request (Required for Local/Intranet Sites)</option>
                  </select>
                </div>
                <div className="form-group">
                  <label>Authentication Type</label>
                  <select id="erp-auth_type" value={profileEditor.auth_type} onChange={handleEditorChange}>
                    <option value="token">API Key & API Secret Token (Recommended)</option>
                    <option value="password">Username & Password Login</option>
                  </select>
                </div>
              </div>

              {profileEditor.auth_type === 'token' ? (
                <div className="form-row">
                  <div className="form-group">
                    <label>API Key <span className="required">*</span></label>
                    <input type="text" id="erp-api_key" value={profileEditor.api_key || ''} onChange={handleEditorChange} required={profileEditor.auth_type === 'token'} />
                  </div>
                  <div className="form-group">
                    <label>API Secret <span className="required">*</span></label>
                    <input type="password" id="erp-api_secret" value={profileEditor.api_secret || ''} onChange={handleEditorChange} required={profileEditor.auth_type === 'token'} />
                  </div>
                </div>
              ) : (
                <div className="form-row">
                  <div className="form-group">
                    <label>Username / Email <span className="required">*</span></label>
                    <input type="text" id="erp-username" value={profileEditor.username || ''} onChange={handleEditorChange} required={profileEditor.auth_type === 'password'} />
                  </div>
                  <div className="form-group">
                    <label>Password <span className="required">*</span></label>
                    <input type="password" id="erp-password" value={profileEditor.password || ''} onChange={handleEditorChange} required={profileEditor.auth_type === 'password'} />
                  </div>
                </div>
              )}

              <div className="form-group">
                <label>Company Name <span className="required">*</span></label>
                <select 
                  id="erp-company" 
                  value={profileEditor.company || ''} 
                  onChange={handleEditorChange} 
                  required 
                >
                  <option value="">-- Select Company --</option>
                  {profileEditor.company && !companiesList.some(c => c.name === profileEditor.company) && (
                    <option value={profileEditor.company}>{profileEditor.company}</option>
                  )}
                  {companiesList.map(c => <option key={c.name} value={c.name}>{c.name}</option>)}
                </select>
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Expense Account (GL) <span className="required">*</span></label>
                  <select 
                    id="erp-expense_account" 
                    value={profileEditor.expense_account || ''} 
                    onChange={handleEditorChange} 
                    required 
                  >
                    <option value="">-- Select Expense Account --</option>
                    {profileEditor.expense_account && !accountsList.some(a => a.name === profileEditor.expense_account) && (
                      <option value={profileEditor.expense_account}>{profileEditor.expense_account}</option>
                    )}
                    {accountsList.map(a => <option key={a.name} value={a.name}>{a.name}</option>)}
                  </select>
                </div>
                <div className="form-group">
                  <label>Cost Center</label>
                  <select 
                    id="erp-cost_center" 
                    value={profileEditor.cost_center || ''} 
                    onChange={handleEditorChange} 
                  >
                    <option value="">-- Select Cost Center --</option>
                    {profileEditor.cost_center && !costCentersList.some(cc => cc.name === profileEditor.cost_center) && (
                      <option value={profileEditor.cost_center}>{profileEditor.cost_center}</option>
                    )}
                    {costCentersList.map(cc => <option key={cc.name} value={cc.name}>{cc.name}</option>)}
                  </select>
                </div>
                <div className="form-group">
                  <label>Default Warehouse (for Receipts)</label>
                  <select 
                    id="erp-warehouse" 
                    value={profileEditor.warehouse || ''} 
                    onChange={handleEditorChange} 
                  >
                    <option value="">-- Select Warehouse --</option>
                    {profileEditor.warehouse && !warehousesList.some(w => w.name === profileEditor.warehouse) && (
                      <option value={profileEditor.warehouse}>{profileEditor.warehouse}</option>
                    )}
                    {warehousesList.map(w => <option key={w.name} value={w.name}>{w.name}</option>)}
                  </select>
                </div>
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Payment Account (Bank/Cash Asset) <span className="required">*</span></label>
                  <select 
                    id="erp-payment_account" 
                    value={profileEditor.payment_account || ''} 
                    onChange={handleEditorChange} 
                    required 
                  >
                    <option value="">-- Select Payment Account --</option>
                    {profileEditor.payment_account && !accountsList.some(a => a.name === profileEditor.payment_account) && (
                      <option value={profileEditor.payment_account}>{profileEditor.payment_account}</option>
                    )}
                    {accountsList.map(a => <option key={a.name} value={a.name}>{a.name}</option>)}
                  </select>
                </div>
                <div className="form-group">
                  <label>Creditors Account (Liability) <span className="required">*</span></label>
                  <select 
                    id="erp-creditors_account" 
                    value={profileEditor.creditors_account || ''} 
                    onChange={handleEditorChange} 
                    required 
                  >
                    <option value="">-- Select Creditors Account --</option>
                    {profileEditor.creditors_account && !accountsList.some(a => a.name === profileEditor.creditors_account) && (
                      <option value={profileEditor.creditors_account}>{profileEditor.creditors_account}</option>
                    )}
                    {accountsList.map(a => <option key={a.name} value={a.name}>{a.name}</option>)}
                  </select>
                </div>
                <div className="form-group">
                  <label>Purchase Taxes Template</label>
                  <select 
                    id="erp-tax_template" 
                    value={profileEditor.tax_template || ''} 
                    onChange={handleEditorChange} 
                  >
                    <option value="">-- No Tax Template --</option>
                    {profileEditor.tax_template && !taxTemplatesList.some(t => t.name === profileEditor.tax_template) && (
                      <option value={profileEditor.tax_template}>{profileEditor.tax_template}</option>
                    )}
                    {taxTemplatesList.map(t => <option key={t.name} value={t.name}>{t.name}</option>)}
                  </select>
                </div>
                <div className="form-group" style={{ gridColumn: '1 / -1', marginTop: '10px' }}>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', margin: 0 }}>
                    <input 
                      type="checkbox" 
                      id="erp-enable_sn_tracking"
                      checked={profileEditor.enable_sn_tracking || false}
                      onChange={(e) => setProfileEditor({...profileEditor, enable_sn_tracking: e.target.checked})}
                      style={{ width: '16px', height: '16px' }}
                    />
                    <span>Enable Serial Number Tracking (Requires manual SN entry or AI extraction for tracking inventory items)</span>
                  </label>
                </div>
              </div>

              <datalist id="erp-accounts-datalist">
                {accountsList.map(a => <option key={a.name} value={a.name}>{a.name}</option>)}
              </datalist>
              <datalist id="erp-costcenters-datalist">
                {costCentersList.map(cc => <option key={cc.name} value={cc.name}>{cc.name}</option>)}
              </datalist>
              <datalist id="erp-items-datalist">
                {itemsList.map(it => <option key={it.name} value={it.item_code}>{it.item_name} ({it.item_code})</option>)}
              </datalist>
              <datalist id="erp-warehouses-datalist">
                {warehousesList.map(w => <option key={w.name} value={w.name}>{w.name}</option>)}
              </datalist>

              <div style={{ display: 'flex', gap: '10px', marginTop: '20px' }}>
                <button type="button" className="btn btn-secondary-outline" onClick={() => testProfileConnection(profileEditor)}>Connect / Test Connection</button>
                <button type="submit" className="btn btn-primary" style={{ flex: 1 }}>Save Profile Settings</button>
              </div>
            </form>
          )}
        </div>
      )}

      {/* Panel 5: AI Parser Settings */}
      {activeTab === 'parser-settings' && (
        <div className="form-card" style={{ maxWidth: '800px', margin: '0 auto' }}>
          <h2>AI Parser Configuration</h2>
          <form onSubmit={(e) => { e.preventDefault(); saveParserSettings(parserSettings); }}>
            <div className="form-group">
              <label>Parser Extraction Engine</label>
              <select value={parserSettings.type} onChange={(e) => setParserSettings({ ...parserSettings, type: e.target.value })}>
                <option value="regex">Rule-Based / Offline Local OCR (Free, No Keys Needed)</option>
                <option value="gemini">Google Gemini 3.6 Flash API (Recommended)</option>
                <option value="openai">OpenAI GPT API</option>
              </select>
            </div>

            {parserSettings.type === 'gemini' && (
              <div className="form-group">
                <label>Google Gemini API Key <span className="required">*</span></label>
                <input type="password" value={parserSettings.geminiKey} onChange={(e) => setParserSettings({ ...parserSettings, geminiKey: e.target.value })} placeholder="AIzaSy..." required />
                <small style={{ color: 'var(--text-muted)' }}>Your API Key is saved locally in your browser storage.</small>
              </div>
            )}

            {parserSettings.type === 'openai' && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
                <div className="form-group">
                  <label>OpenAI API Key <span className="required">*</span></label>
                  <input type="password" value={parserSettings.openaiKey} onChange={(e) => setParserSettings({ ...parserSettings, openaiKey: e.target.value })} placeholder="sk-..." required />
                </div>
                <div className="form-group">
                  <label>OpenAI Chat Model</label>
                  <select value={parserSettings.openaiModel} onChange={(e) => setParserSettings({ ...parserSettings, openaiModel: e.target.value })}>
                    <option value="gpt-4o-mini">gpt-4o-mini (Cost-effective, highly accurate)</option>
                    <option value="gpt-4o">gpt-4o (Standard advanced model)</option>
                  </select>
                </div>
              </div>
            )}

            <button type="submit" className="btn btn-primary" style={{ marginTop: '16px' }}>Save Parser Settings</button>
          </form>
        </div>
      )}

      {/* Panel: Rebate Reconciliation Module */}
      {activeTab === 'reconciliation' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
          <div className="form-card">
            <h2>📈 Quarterly Rebate Reconciliation Module</h2>
            <p style={{ color: 'var(--text-muted)', fontSize: '13px', marginBottom: '16px' }}>
              Upload a payment list document or statement image. Gemini AI Vision will extract invoice numbers and reconcile rebates stored in Neon DB for active profile <strong>{activeProfileName || 'Default'}</strong>.
            </p>

            <div 
              className="drop-zone"
              onClick={() => document.getElementById('recon-file-input')?.click()}
              style={{ cursor: 'pointer', textAlign: 'center', padding: '30px', border: '2px dashed rgba(56, 189, 248, 0.3)', borderRadius: 'var(--radius-md)' }}
            >
              <div className="drop-zone-icon">📄</div>
              <p className="drop-zone-text">Upload Payment List Document or Image</p>
              <small style={{ color: reconFileColor, fontWeight: '500' }}>{reconFileStatus}</small>
              <input 
                id="recon-file-input"
                type="file" 
                accept="image/*,application/pdf"
                onChange={handleReconFileUpload}
                style={{ display: 'none' }} 
              />
            </div>
          </div>

          <div className="analytics-grid">
            <div className="analytic-card">
              <span className="analytic-title">Total Invoices Matched</span>
              <span className="analytic-val">{reconInvoices.length}</span>
              <span className="analytic-sub">Extracted via Gemini Vision</span>
            </div>
            <div className="analytic-card">
              <span className="analytic-title">Reconciled Line Items</span>
              <span className="analytic-val">{reconRecords.length}</span>
              <span className="analytic-sub">Neon DB Records</span>
            </div>
            <div className="analytic-card">
              <span className="analytic-title">Total Quarterly Rebate</span>
              <span className="analytic-val">
                ${reconRecords.reduce((sum, r) => sum + parseFloat(r.rebate_amount || 0), 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </span>
              <span className="analytic-sub">Active Profile: {activeProfileName}</span>
            </div>
          </div>

          <div className="form-card">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <div>
                <h2>Reconciled Rebate Records Breakdown</h2>
                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                  Showing records matching payment list invoice numbers.
                </span>
              </div>
              <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                <input 
                  type="text" 
                  placeholder="Filter invoice or model..." 
                  value={reconSearch}
                  onChange={(e) => setReconSearch(e.target.value)}
                  style={{ padding: '6px 12px', fontSize: '12px', width: '200px' }}
                />
                <button type="button" className="btn" style={{ backgroundColor: 'var(--bg-lighter)', border: '1px solid var(--border)', color: 'var(--text)' }} onClick={loadAllRebateRecords} disabled={reconLoading}>[+] Load All Records</button> <button 
                  type="button" 
                  className="btn btn-primary"
                  onClick={exportReconToExcel}
                  disabled={reconRecords.length === 0}
                >
                  📥 Export to Excel (.xlsx)
                </button>
              </div>
            </div>

            <div className="table-container">
              <table className="items-table">
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>Doc Name</th>
                    <th>Supplier Inv.</th>
                    <th>Model Code</th>
                    <th>Qty</th>
                    <th>Rate</th>
                    <th>Real Rate</th>
                    <th>Rebate %</th>
                    <th>Rebate Amount ($)</th>
                    <th>Profile</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {reconRecords
                    .filter(r => {
                      const q = reconSearch.toLowerCase();
                      return (r.invoice_no || '').toLowerCase().includes(q) || (r.supplier_invoice_no || '').toLowerCase().includes(q) || (r.item_model_code || '').toLowerCase().includes(q);
                    })
                    .map((r, idx) => (
                      <tr key={idx}>
                        <td>{r.date ? new Date(r.date).toISOString().split('T')[0] : ''}</td>
                        <td><strong>{r.invoice_no}</strong></td>
                        <td>{r.supplier_invoice_no ? r.supplier_invoice_no : <span style={{color:'var(--text-muted)'}}>N/A</span>}</td>
                        <td><code>{r.item_model_code}</code></td>
                        <td>{parseFloat(r.quantity)}</td>
                        <td>${parseFloat(r.rate).toFixed(2)}</td>
                        <td>${parseFloat(r.real_rate).toFixed(2)}</td>
                        <td>
                          <span className="badge" style={{ backgroundColor: 'rgba(56, 189, 248, 0.15)', color: 'var(--color-secondary)' }}>
                            {(parseFloat(r.rebate_percentage) * 100).toFixed(1)}%
                          </span>
                        </td>
                        <td>
                          <strong style={{ color: 'var(--status-success)' }}>
                            ${parseFloat(r.rebate_amount).toFixed(2)}
                          </strong>
                        </td>
                        <td>{r.company_profile}</td>
                        <td>
                          <button 
                            type="button" 
                            style={{ background: 'transparent', border: 'none', color: 'var(--status-error)', cursor: 'pointer', fontSize: '14px', fontWeight: 'bold' }}
                            onClick={() => deleteRebateRecord(r.id)}
                            title="Delete Record"
                          >
                            [X]
                          </button>
                        </td>
                      </tr>
                    ))}
                  {reconRecords.length === 0 && (
                    <tr>
                      <td colSpan="9" style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '20px' }}>
                        No reconciled rebate records found. Upload a payment list document to perform quarterly reconciliation.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Panel: Rebate Exceptions Manager */}
      {activeTab === 'rebate-settings' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px', maxWidth: '1000px', margin: '0 auto' }}>
          <div className="form-card">
            <h2>⚙️ Rebate Exception List Manager (8% Rule)</h2>
            <p style={{ color: 'var(--text-muted)', fontSize: '13px', marginBottom: '16px' }}>
              Upload an Excel (.xlsx) or CSV file containing item model codes eligible for the <strong>8% rebate rule</strong> (DS, iDS, CS models). Uploading will replace the exception list in Neon DB.
            </p>

            <div 
              className="drop-zone"
              onClick={() => document.getElementById('exception-file-input')?.click()}
              style={{ cursor: 'pointer', textAlign: 'center', padding: '30px', border: '2px dashed rgba(56, 189, 248, 0.3)', borderRadius: 'var(--radius-md)' }}
            >
              <div className="drop-zone-icon">📊</div>
              <p className="drop-zone-text">Select Excel (.xlsx) or CSV File</p>
              <small style={{ color: exceptionUploadColor, fontWeight: '500' }}>{exceptionUploadStatus}</small>
              <input 
                id="exception-file-input"
                type="file" 
                accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                onChange={handleExceptionListUpload}
                style={{ display: 'none' }} 
              />
            </div>
          </div>

          <div className="form-card">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <div>
                <h2>Active Masterlist Models ({rebateExceptionList.length})</h2>
                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                  Filtered masterlist containing DS, iDS, CS (8%), RG (5%), and VS (10%) models. For DS/iDS/CS items NOT in this list, rebate defaults to 16%.
                </span>
              </div>
              <input 
                type="text" 
                placeholder="Search model code..." 
                value={exceptionSearch}
                onChange={(e) => setExceptionSearch(e.target.value)}
                style={{ padding: '6px 12px', fontSize: '12px', width: '220px' }}
              />
            </div>

            <div className="table-container" style={{ maxHeight: '400px', overflowY: 'auto' }}>
              <table className="items-table">
                <thead>
                  <tr>
                    <th style={{ width: '10%' }}>#</th>
                    <th style={{ width: '60%' }}>Item Model Code</th>
                    <th style={{ width: '30%' }}>Rebate Rate Rule</th>
                  </tr>
                </thead>
                <tbody>
                  {rebateExceptionList
                    .filter(m => String(m).toLowerCase().includes(exceptionSearch.toLowerCase()))
                    .map((itemCode, index) => {
                      const upperCode = String(itemCode || '').toUpperCase().trim();
                      let rateLabel = '0.0%';
                      let badgeStyle = { backgroundColor: 'rgba(255, 255, 255, 0.05)', color: 'var(--text-muted)' };

                      if (upperCode.startsWith('RG')) {
                        rateLabel = '5.0% (RG Rule)';
                        badgeStyle = { backgroundColor: 'rgba(56, 189, 248, 0.15)', color: 'var(--color-secondary)' };
                      } else if (upperCode.startsWith('VS')) {
                        rateLabel = '10.0% (VS Rule)';
                        badgeStyle = { backgroundColor: 'rgba(168, 85, 247, 0.15)', color: '#c084fc' };
                      } else if (upperCode.startsWith('DS') || upperCode.startsWith('IDS') || upperCode.startsWith('CS')) {
                        rateLabel = '8.0% (Masterlist Exception)';
                        badgeStyle = { backgroundColor: 'rgba(34, 197, 94, 0.15)', color: 'var(--status-success)' };
                      }

                      return (
                        <tr key={index}>
                          <td>{index + 1}</td>
                          <td><code>{itemCode}</code></td>
                          <td>
                            <span className="badge" style={badgeStyle}>
                              {rateLabel}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  {rebateExceptionList.length === 0 && (
                    <tr>
                      <td colSpan="3" style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '20px' }}>
                        No valid exception models currently in list. Upload an Excel file above to populate.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

