/* eslint-disable no-throw-literal */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { IRequiredFields } from "../interfaces/IGlobalInterfaces";

export class BaselineHelper {

    
  public static checkRequiredFields = (
    currentStates: any,
    requiredFields: IRequiredFields[]
  ) => {
    const errorTargets: IRequiredFields[] = [];
    requiredFields.forEach((i) => {
      const stateCheck = currentStates[i.stateName];
      if (typeof currentStates[i.stateName] === "object" && i.stateName.toLocaleLowerCase().indexOf('object') > 0) {
        stateCheck.map((po:any,index:number)=>{
          i.fields?.map((flds)=>{
            const objectStateCheck =po[flds];
            if (objectStateCheck === undefined || objectStateCheck === "" || objectStateCheck === null ||  objectStateCheck.length === 0) {
              errorTargets.push({ stateName: i.stateName, fieldName: flds , index:index});
            }
          })
        })
        
      }
      else
        if ( stateCheck === undefined ||stateCheck === "" || stateCheck === null || stateCheck.length === 0 ) {
          errorTargets.push({ stateName: i.stateName, fieldName: i.fieldName });
        }
    });
    if (errorTargets.length > 1) {
      throw {
        targetFields: errorTargets,
        message: "The below fields are required.",
      };
    } else if (errorTargets.length === 1) {
      throw {
        targetFields: errorTargets,
        message: 'The field "' + errorTargets[0].fieldName + '" is required',
      };
    }
  };

}
