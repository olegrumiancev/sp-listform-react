const escapeChars = { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' };

export const formHelper = {
  unescapeHTML: (str) => {
    return str.replace(/\&([^;]+);/g, (entity, entityCode) => {
      let match;

      if (entityCode in escapeChars) {
        return escapeChars[entityCode];
      // tslint:disable-next-line:no-conditional-assignment
      } else if (match = entityCode.match(/^#x([\da-fA-F]+)$/)) {
        return String.fromCharCode(parseInt(match[1], 16));
      // tslint:disable-next-line:no-conditional-assignment
      } else if (match = entityCode.match(/^#(\d+)$/)) {
        return String.fromCharCode(~~match[1]);
      } else {
        return entity;
      }
    });
  },
  handleError: (msg: string) => {
    console.error(msg);
  }
};
