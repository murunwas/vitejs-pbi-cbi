import * as pbi from 'powerbi-client';
import { EMBED_URL, REPORT_ID, TOKEN } from './constants';
const permissions = pbi.models.Permissions.All;

let loadedResolve,
  reportLoaded = new Promise((res, rej) => {
    loadedResolve = res;
  });
let renderedResolve,
  reportRendered = new Promise((res, rej) => {
    renderedResolve = res;
  });

export function initPBI() {
  const config = {
    type: 'report',
    tokenType: pbi.models.TokenType.Embed,
    accessToken: TOKEN,
    embedUrl: EMBED_URL,
    id: REPORT_ID,
    pageView: 'fitToWidth',
    permissions: permissions,
    settings: {
      panes: {
        filters: {
          visible: false,
        },
        pageNavigation: {
          visible: false,
        },
      },
    },
  };

  let powerbi = new pbi.service.Service(
    pbi.factories.hpmFactory,
    pbi.factories.wpmpFactory,
    pbi.factories.routerFactory
  );
  const dashboardContainer = document.getElementById('container');
  const dashboard = powerbi.embed(dashboardContainer, config);

  dashboard.off('loaded');
  dashboard.on('loaded', function () {
    loadedResolve();
    dashboard.off('loaded');
  });
  dashboard.off('rendered');

  // report.on will add an event handler
  dashboard.on('rendered', function () {
    renderedResolve();
    dashboard.off('rendered');
  });

  dashboard.on('error', function () {
    this.dashboard.off('error');
  });

  return {
    async print() {
      try {
        await dashboard.print();
      } catch (errors) {
        console.log(errors);
      }
    },
  };
}
