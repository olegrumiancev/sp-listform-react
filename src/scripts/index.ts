// require('core-js/es6');
// import 'core-js/es6';
import 'core-js/shim';
import * as ReactDOM from 'react-dom';
import root from './root';

const container: HTMLElement = document.getElementById('listform-cewp-container');

SP.SOD.executeOrDelayUntilScriptLoaded(() => {
  ReactDOM.render(root, container);
}, 'SP.js');
