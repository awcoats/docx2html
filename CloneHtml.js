/**
 * Serializes the HTML contents of an element, baking in the live state
 * of checkboxes, radios, text inputs, textareas, and selects — even if
 * those controls have no id or name.
 *
 * @param {string} elementId - id of the container element to serialize.
 * @returns {string} HTML string reflecting current form control state.
 */
function serializeElementHTML(elementId) {
  const container = document.getElementById(elementId);
  if (!container) {
    throw new Error(`No element found with id "${elementId}"`);
  }

  const clone = container.cloneNode(true);

  const originalControls = container.querySelectorAll('input, textarea, select');
  const cloneControls = clone.querySelectorAll('input, textarea, select');

  originalControls.forEach((orig, i) => {
    const copy = cloneControls[i];
    const tag = orig.tagName;

    if (tag === 'INPUT') {
      const type = (orig.getAttribute('type') || 'text').toLowerCase();
      if (type === 'checkbox' || type === 'radio') {
        if (orig.checked) {
          copy.setAttribute('checked', '');
        } else {
          copy.removeAttribute('checked');
        }
      } else {
        copy.setAttribute('value', orig.value);
      }
    } else if (tag === 'TEXTAREA') {
      copy.textContent = orig.value;
    } else if (tag === 'SELECT') {
      Array.from(copy.options).forEach((opt, idx) => {
        if (orig.options[idx].selected) {
          opt.setAttribute('selected', '');
        } else {
          opt.removeAttribute('selected');
        }
      });
    }
  });

  return clone.innerHTML;
}