/**
 * loads and decorates the block
 * @param {Element} block The block element
 */
export default function decorate(block) {
  const title = document.createElement('p');
  title.className = 'aram-test-title';
  title.textContent = 'Aram Test';
  block.replaceChildren(title);
}
