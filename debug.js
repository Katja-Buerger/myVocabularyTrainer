const debug = (...args) => {
  const uploadMessage = document.getElementById('uploadMessage');
  if (uploadMessage) {
    const message = document.createElement('div');
    message.style.fontSize = '0.8rem';
    message.style.color = '#555';
    message.textContent = args.map((value) => {
      try {
        if (typeof value === 'object') {
          return JSON.stringify(value);
        }
        return String(value);
      } catch (error) {
        return '[unserializable]';
      }
    }).join(' ');
    uploadMessage.appendChild(message);
  }
};
debug('script loaded');
