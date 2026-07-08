const STORAGE_KEY = 'sonhoabinh-production-v1';
const STORAGE_QUOTA_MESSAGE = 'Bộ nhớ trình duyệt đã đầy. Vui lòng Reset dữ liệu demo hoặc chuyển sang database.';

function isQuotaExceededError(error) {
  return error?.name === 'QuotaExceededError'
    || error?.name === 'NS_ERROR_DOM_QUOTA_REACHED'
    || error?.code === 22
    || error?.code === 1014;
}

export function loadProductionData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export function saveProductionData(data) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch (error) {
    if (isQuotaExceededError(error)) console.warn(STORAGE_QUOTA_MESSAGE);
    else console.warn('Không thể lưu dữ liệu production vào localStorage.', error);
  }
}
