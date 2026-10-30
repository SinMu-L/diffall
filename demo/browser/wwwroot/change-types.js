// Keep preview colors and the change-list categories aligned across formats.
export const changeGroup = change => change.aspect === 'existence' ? 'existence'
  : ['style', 'geometry', 'order', 'parent'].includes(change.aspect) ? 'layout' : 'content';
