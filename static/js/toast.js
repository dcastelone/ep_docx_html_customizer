/**
 * Toast notification system for ep_docx_html_customizer
 * Shows progress updates for image uploads during import and clipboard operations
 */

class DocxToastManager {
  constructor() {
    this.container = null;
    this.toasts = new Map();
    this.ensureContainer();
  }

  ensureContainer() {
    if (!this.container) {
      this.container = document.createElement('div');
      this.container.className = 'docx-toast-container';
      document.body.appendChild(this.container);
    }
  }

  createToast(id, title, message = '', type = 'default') {
    const toast = document.createElement('div');
    toast.className = `docx-toast ${type}`;
    toast.innerHTML = `
      <button class="docx-toast-close" onclick="docxToast.removeToast('${id}')">&times;</button>
      <div class="docx-toast-title">${title}</div>
      ${message ? `<div class="docx-toast-message">${message}</div>` : ''}
      <div class="docx-toast-progress"></div>
      <div class="docx-toast-progress-bar">
        <div class="docx-toast-progress-fill"></div>
      </div>
    `;

    this.container.appendChild(toast);
    this.toasts.set(id, toast);

    // Show with animation
    setTimeout(() => toast.classList.add('show'), 100);

    return toast;
  }

  updateProgress(id, current, total, message = '') {
    const toast = this.toasts.get(id);
    if (!toast) return;

    const progressText = toast.querySelector('.docx-toast-progress');
    const progressFill = toast.querySelector('.docx-toast-progress-fill');
    
    const percentage = total > 0 ? Math.round((current / total) * 100) : 0;
    
    progressText.textContent = message || `${current} of ${total} completed`;
    progressFill.style.width = `${percentage}%`;
  }

  showToast(id, title, message = '', type = 'default', autoDismiss = 0) {
    this.removeToast(id); // Remove existing toast with same ID
    const toast = this.createToast(id, title, message, type);

    if (autoDismiss > 0) {
      setTimeout(() => this.removeToast(id), autoDismiss);
    }

    return toast;
  }

  updateToast(id, title, message = '', type = null) {
    const toast = this.toasts.get(id);
    if (!toast) return;

    const titleEl = toast.querySelector('.docx-toast-title');
    const messageEl = toast.querySelector('.docx-toast-message');
    
    if (titleEl) titleEl.textContent = title;
    if (messageEl) {
      if (message) {
        messageEl.textContent = message;
        messageEl.style.display = 'block';
      } else {
        messageEl.style.display = 'none';
      }
    }

    if (type) {
      toast.className = `docx-toast ${type} show`;
    }
  }

  removeToast(id) {
    const toast = this.toasts.get(id);
    if (!toast) return;

    toast.classList.remove('show');
    setTimeout(() => {
      if (toast.parentNode) {
        toast.parentNode.removeChild(toast);
      }
      this.toasts.delete(id);
    }, 300);
  }

  showImageUploadProgress(id, title = 'Uploading Images') {
    return this.showToast(id, title);
  }

  updateImageUploadProgress(id, current, total, details = '') {
    this.updateProgress(id, current, total, `${current} of ${total} images uploaded${details ? ` (${details})` : ''}`);
  }

  completeImageUpload(id, successCount, totalCount, errorCount = 0) {
    if (errorCount > 0) {
      this.updateToast(id, 'Upload Complete with Errors', 
        `${successCount} uploaded, ${errorCount} failed`, 'warning');
    } else {
      this.updateToast(id, 'Upload Complete', 
        `${successCount} images uploaded successfully`, 'success');
    }
    
    // Hide progress bar when complete
    const toast = this.toasts.get(id);
    if (toast) {
      const progressBar = toast.querySelector('.docx-toast-progress-bar');
      if (progressBar) progressBar.style.display = 'none';
    }

    // Auto-dismiss after 4 seconds
    setTimeout(() => this.removeToast(id), 4000);
  }

  showError(id, title, message, autoDismiss = 5000) {
    this.showToast(id, title, message, 'error', autoDismiss);
  }
}

// Global instance
window.docxToast = new DocxToastManager(); 