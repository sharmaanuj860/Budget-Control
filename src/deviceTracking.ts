// Device Hardware and Public IP / Location Profiling Engine
// Captures detailed device specifications and IP geolocation telemetry to detect unauthorized access

export interface DeviceHardwareProfile {
  screenResolution: string;      // e.g. "1920 x 1080"
  availableResolution: string; // e.g. "1920 x 1040"
  colorDepth: string;          // e.g. "24-bit"
  pixelRatio: string;          // e.g. "2x DPR"
  orientation: string;         // e.g. "Landscape (landscape-primary)"
  viewportSize: string;        // e.g. "1440 x 820"
  hardwareConcurrency: string; // e.g. "8 Cores"
  deviceMemory: string;        // e.g. "8 GB RAM"
  maxTouchPoints: number;
  touchSupport: string;        // e.g. "Touch Screen" or "Non-touch"
  platform: string;            // e.g. "Win32" / "MacIntel" / "Linux x86_64"
  os: string;                  // e.g. "Windows 10/11", "macOS", "Android 14", "iOS"
  browser: string;             // e.g. "Chrome", "Firefox", "Safari", "Edge"
  browserVersion: string;      // e.g. "124.0.0.0"
  deviceType: 'Desktop / PC' | 'Mobile Phone' | 'Tablet';
  timeZone: string;            // e.g. "Asia/Kolkata"
  languages: string;           // e.g. "en-US, en"
  gpuRenderer: string;         // e.g. "Intel Iris Xe Graphics" or "WebGL Enabled"
  hardwareSummary: string;     // Compact string: "1920x1080 • 8 Cores • 8 GB RAM • 2x DPR"
  summary: string;             // Overall summary: "Windows 10/11 (Chrome 124.0) • Desktop • 1920x1080"
  userAgent: string;
}

export interface NetworkLocationProfile {
  ip: string;
  city: string;
  region: string;
  country: string;
  countryCode: string;
  postal: string;
  isp: string;
  org: string;
  locationString: string;
  timezone: string;
  isVpnOrProxy?: boolean;
}

// In-memory cache for IP & Geolocation
let cachedNetworkProfile: NetworkLocationProfile | null = null;
let isFetchingNetwork = false;
const fetchListeners: Array<(profile: NetworkLocationProfile) => void> = [];

export function getGpuRendererSafe(): string {
  if (typeof window === 'undefined' || typeof document === 'undefined') return 'N/A';
  try {
    const canvas = document.createElement('canvas');
    const gl = canvas.getContext('webgl') || canvas.getContext('experimental-webgl');
    if (gl && 'getExtension' in gl) {
      const debugInfo = (gl as any).getExtension('WEBGL_debug_renderer_info');
      if (debugInfo) {
        const renderer = (gl as any).getParameter(debugInfo.UNMASKED_RENDERER_WEBGL);
        if (renderer) return String(renderer).replace(/ANGLE \((.*)\)/, '$1').trim();
      }
    }
    return gl ? 'WebGL Active' : 'N/A';
  } catch (_) {
    return 'N/A';
  }
}

export function getDeviceHardwareProfile(): DeviceHardwareProfile {
  if (typeof window === 'undefined' || !window.navigator) {
    return {
      screenResolution: 'N/A',
      availableResolution: 'N/A',
      colorDepth: 'N/A',
      pixelRatio: '1x',
      orientation: 'N/A',
      viewportSize: 'N/A',
      hardwareConcurrency: 'N/A',
      deviceMemory: 'N/A',
      maxTouchPoints: 0,
      touchSupport: 'Non-touch',
      platform: 'Server',
      os: 'Unknown OS',
      browser: 'Web Client',
      browserVersion: 'N/A',
      deviceType: 'Desktop / PC',
      timeZone: 'UTC',
      languages: 'en',
      gpuRenderer: 'N/A',
      hardwareSummary: 'Standard Web Client',
      summary: 'Web Client',
      userAgent: ''
    };
  }

  const ua = window.navigator.userAgent || '';
  const screen = window.screen || {} as Screen;
  
  // OS Detection
  let os = 'Windows / PC';
  if (/windows phone/i.test(ua)) os = 'Windows Phone';
  else if (/win(dows )?nt 10\.0/i.test(ua)) os = 'Windows 10/11';
  else if (/win(dows )?nt 6\.3/i.test(ua)) os = 'Windows 8.1';
  else if (/win(dows )?nt 6\.2/i.test(ua)) os = 'Windows 8';
  else if (/win(dows )?nt 6\.1/i.test(ua)) os = 'Windows 7';
  else if (/android/i.test(ua)) {
    const androidVer = ua.match(/android\s([0-9\.]+)/i);
    os = androidVer ? `Android ${androidVer[1]}` : 'Android';
  } else if (/ipad/i.test(ua)) {
    os = 'iPadOS / iOS';
  } else if (/iphone|ipod/i.test(ua)) {
    const iosVer = ua.match(/os\s([0-9\_]+)/i);
    os = iosVer ? `iOS ${iosVer[1].replace(/_/g, '.')}` : 'iPhone iOS';
  } else if (/macintosh|mac os x/i.test(ua)) {
    const macVer = ua.match(/mac os x\s([0-9\_\.]+)/i);
    os = macVer ? `macOS ${macVer[1].replace(/_/g, '.')}` : 'macOS';
  } else if (/cros/i.test(ua)) {
    os = 'ChromeOS';
  } else if (/linux/i.test(ua)) {
    os = 'Linux';
  }

  // Browser Detection & Version
  let browser = 'Browser';
  let browserVersion = '';
  
  if (/edg\/([0-9\.]+)/i.test(ua)) {
    browser = 'Microsoft Edge';
    browserVersion = ua.match(/edg\/([0-9\.]+)/i)?.[1] || '';
  } else if (/opr\/([0-9\.]+)|opera\/([0-9\.]+)/i.test(ua)) {
    browser = 'Opera';
    browserVersion = ua.match(/opr\/([0-9\.]+)|opera\/([0-9\.]+)/i)?.[1] || '';
  } else if (/samsungbrowser\/([0-9\.]+)/i.test(ua)) {
    browser = 'Samsung Internet';
    browserVersion = ua.match(/samsungbrowser\/([0-9\.]+)/i)?.[1] || '';
  } else if (/chrome|crios/i.test(ua)) {
    browser = 'Chrome';
    browserVersion = ua.match(/(?:chrome|crios)\/([0-9\.]+)/i)?.[1] || '';
  } else if (/firefox|fxios/i.test(ua)) {
    browser = 'Firefox';
    browserVersion = ua.match(/(?:firefox|fxios)\/([0-9\.]+)/i)?.[1] || '';
  } else if (/safari/i.test(ua) && !/chrome|crios|android/i.test(ua)) {
    browser = 'Safari';
    browserVersion = ua.match(/version\/([0-9\.]+)/i)?.[1] || '';
  }

  // Device Classification
  const isTablet = /(ipad|tablet|(android(?!.*mobile))|(windows(?!.*phone)(.*touch))|kindle|playbook|silk)/i.test(ua);
  const isMobile = !isTablet && /(mobi|ipod|phone|blackberry|opera mini|fennec|minimo|symbian|psp|nintendo)/i.test(ua);
  const deviceType: 'Desktop / PC' | 'Mobile Phone' | 'Tablet' = isTablet ? 'Tablet' : (isMobile ? 'Mobile Phone' : 'Desktop / PC');

  // Screen Hardware specs
  const screenW = screen.width || window.innerWidth || 0;
  const screenH = screen.height || window.innerHeight || 0;
  const screenResolution = `${screenW} x ${screenH}`;
  const availW = screen.availWidth || screenW;
  const availH = screen.availHeight || screenH;
  const availableResolution = `${availW} x ${availH}`;
  const colorDepth = screen.colorDepth ? `${screen.colorDepth}-bit` : '24-bit';
  const dpr = typeof window.devicePixelRatio === 'number' ? `${window.devicePixelRatio.toFixed(1)}x` : '1.0x';
  
  let orientation = 'Landscape';
  if (screen.orientation?.type) {
    orientation = screen.orientation.type.includes('portrait') ? 'Portrait' : 'Landscape';
  } else if (window.innerHeight > window.innerWidth) {
    orientation = 'Portrait';
  }

  const viewportSize = `${window.innerWidth || 0} x ${window.innerHeight || 0}`;
  
  // CPU cores & RAM
  const hwConcurrency = (window.navigator as any).hardwareConcurrency ? `${(window.navigator as any).hardwareConcurrency} Cores` : 'N/A';
  const devMem = (window.navigator as any).deviceMemory ? `${(window.navigator as any).deviceMemory} GB RAM` : 'N/A';
  const touchPoints = window.navigator.maxTouchPoints || 0;
  const touchSupport = touchPoints > 0 ? `Touchscreen (${touchPoints} pts)` : 'Non-touch';
  const platform = window.navigator.platform || 'Unknown';
  
  let timeZone = 'Asia/Kolkata';
  try {
    timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || 'Asia/Kolkata';
  } catch (_) {}

  const languages = window.navigator.languages ? window.navigator.languages.join(', ') : (window.navigator.language || 'en');
  const gpuRenderer = getGpuRendererSafe();

  const hardwareParts = [screenResolution, `${dpr} DPR`];
  if (hwConcurrency !== 'N/A') hardwareParts.push(hwConcurrency);
  if (devMem !== 'N/A') hardwareParts.push(devMem);
  if (touchPoints > 0) hardwareParts.push('Touch');
  const hardwareSummary = hardwareParts.join(' • ');

  const summary = `${os} (${browser}${browserVersion ? ' ' + browserVersion.split('.')[0] : ''}) • ${deviceType} • ${screenResolution}`;

  return {
    screenResolution,
    availableResolution,
    colorDepth,
    pixelRatio: dpr,
    orientation,
    viewportSize,
    hardwareConcurrency: hwConcurrency,
    deviceMemory: devMem,
    maxTouchPoints: touchPoints,
    touchSupport,
    platform,
    os,
    browser,
    browserVersion,
    deviceType,
    timeZone,
    languages,
    gpuRenderer,
    hardwareSummary,
    summary,
    userAgent: ua
  };
}

export async function fetchClientNetworkLocation(): Promise<NetworkLocationProfile> {
  if (cachedNetworkProfile && cachedNetworkProfile.ip && cachedNetworkProfile.ip !== 'Detecting...' && cachedNetworkProfile.ip !== 'Direct IP') {
    return cachedNetworkProfile;
  }

  if (isFetchingNetwork) {
    return new Promise((resolve) => {
      fetchListeners.push(resolve);
    });
  }

  isFetchingNetwork = true;

  const defaultProfile: NetworkLocationProfile = {
    ip: 'Direct IP',
    city: 'Location Detected',
    region: 'Himachal Pradesh',
    country: 'India',
    countryCode: 'IN',
    postal: '',
    isp: 'Direct Connection',
    org: 'Government / State Network',
    locationString: 'Himachal Pradesh, India',
    timezone: 'Asia/Kolkata'
  };

  // Provider 1: ipapi.co
  try {
    const res = await fetch('https://ipapi.co/json/', { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      if (data.ip) {
        const parts = [data.city, data.region, data.country_name || data.country_code].filter(Boolean);
        const profile: NetworkLocationProfile = {
          ip: data.ip,
          city: data.city || 'Unknown City',
          region: data.region || 'Himachal Pradesh',
          country: data.country_name || 'India',
          countryCode: data.country_code || 'IN',
          postal: data.postal || '',
          isp: data.org || data.asn || 'Internet Service Provider',
          org: data.org || '',
          locationString: parts.length > 0 ? parts.join(', ') : 'India',
          timezone: data.timezone || 'Asia/Kolkata'
        };
        cachedNetworkProfile = profile;
        isFetchingNetwork = false;
        fetchListeners.forEach(cb => cb(profile));
        fetchListeners.length = 0;
        return profile;
      }
    }
  } catch (_) {}

  // Provider 2: ipwho.is
  try {
    const res = await fetch('https://ipwho.is/', { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      if (data.success !== false && data.ip) {
        const parts = [data.city, data.region, data.country].filter(Boolean);
        const profile: NetworkLocationProfile = {
          ip: data.ip,
          city: data.city || 'Unknown City',
          region: data.region || '',
          country: data.country || 'India',
          countryCode: data.country_code || 'IN',
          postal: data.postal || '',
          isp: data.connection?.isp || data.connection?.org || 'Internet Service Provider',
          org: data.connection?.org || '',
          locationString: parts.length > 0 ? parts.join(', ') : 'India',
          timezone: data.timezone?.id || 'Asia/Kolkata'
        };
        cachedNetworkProfile = profile;
        isFetchingNetwork = false;
        fetchListeners.forEach(cb => cb(profile));
        fetchListeners.length = 0;
        return profile;
      }
    }
  } catch (_) {}

  // Provider 3: freeipapi.com
  try {
    const res = await fetch('https://freeipapi.com/api/json', { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      if (data.ipAddress) {
        const parts = [data.cityName, data.regionName, data.countryName].filter(Boolean);
        const profile: NetworkLocationProfile = {
          ip: data.ipAddress,
          city: data.cityName || 'Unknown City',
          region: data.regionName || '',
          country: data.countryName || 'India',
          countryCode: data.countryCode || 'IN',
          postal: data.zipCode || '',
          isp: 'Internet Service Provider',
          org: '',
          locationString: parts.length > 0 ? parts.join(', ') : 'India',
          timezone: data.timeZones?.[0] || 'Asia/Kolkata'
        };
        cachedNetworkProfile = profile;
        isFetchingNetwork = false;
        fetchListeners.forEach(cb => cb(profile));
        fetchListeners.length = 0;
        return profile;
      }
    }
  } catch (_) {}

  // Provider 4: api.ipify.org fallback for IP only
  try {
    const res = await fetch('https://api.ipify.org?format=json', { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      if (data.ip) {
        const profile: NetworkLocationProfile = {
          ...defaultProfile,
          ip: data.ip,
          locationString: 'India (IP Verified)'
        };
        cachedNetworkProfile = profile;
        isFetchingNetwork = false;
        fetchListeners.forEach(cb => cb(profile));
        fetchListeners.length = 0;
        return profile;
      }
    }
  } catch (_) {}

  cachedNetworkProfile = defaultProfile;
  isFetchingNetwork = false;
  fetchListeners.forEach(cb => cb(defaultProfile));
  fetchListeners.length = 0;
  return defaultProfile;
}

export function getCachedNetworkProfile(): NetworkLocationProfile | null {
  return cachedNetworkProfile;
}
