import * as ReactDOM from 'react-dom';
import root from './root';

const container: HTMLElement = document.getElementById('listform-cewp-container');

SP.SOD.executeOrDelayUntilScriptLoaded(() => {
  ReactDOM.render(root, container);
  console.log('loaded');
}, 'SP.js');
