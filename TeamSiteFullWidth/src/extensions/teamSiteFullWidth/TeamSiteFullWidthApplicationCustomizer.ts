import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
export interface ITeamSiteFullWidthApplicationCustomizerProperties {
}

require("./css/customStyles.module.scss");

export default class TeamSiteFullWidthApplicationCustomizer extends BaseApplicationCustomizer<ITeamSiteFullWidthApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // tslint:disable-next-line:no-function-expression
      if (typeof (Event) === 'function') {
        window.dispatchEvent(new Event('resize'));
      } else {
        var resizeEvent = window.document.createEvent('UIEvents');
        resizeEvent.initUIEvent('resize', true, false, window, 0);
        window.dispatchEvent(resizeEvent);
      }
    return Promise.resolve();
  }
}
