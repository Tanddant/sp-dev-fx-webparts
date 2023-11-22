import { MSGraphClientV3 } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import {
  IProfileCardPropertiesResults
} from "../Entities/IProfileCardPropertiesResults";
import { IProfileCardProperty } from "../Entities/IProfileCardProperty";
import {
  IUpdateProfileCardProperty
} from "../Entities/IUpdateProfileCardProperty";

const ADMIN_ROLETEMPLATE_ID = "62e90394-69f5-4237-9190-012177145e10";

export const useProfileCardProperties =  () => {
  // Get List of Properties
  const getProfileCardProperties = async (
    msGraphClient: MSGraphClientV3,
  ): Promise<IProfileCardProperty[]> => {

    const _profileProperties: IProfileCardPropertiesResults = await msGraphClient
      .api(`/admin/people/profileCardProperties`)
      .orderby("directoryPropertyName")
      .get();



    return _profileProperties.value;
  };

  // Add Property
  const newProfileCardProperty = async (
    msGraphClient: MSGraphClientV3,
    profileCardProperties: IProfileCardProperty
  ):Promise<IProfileCardProperty> => {

    const _profileProperty: IProfileCardProperty = await msGraphClient
      .api(`/admin/people/profileCardProperties`)
      .post(profileCardProperties);

      return _profileProperty;
  };

  // Update Profile Card Property
  const updateProfileCardProperty = async (
    msGraphClient: MSGraphClientV3,
    profileCardProperties: IProfileCardProperty,
    WebPartContext: WebPartContext
  ): Promise<IProfileCardProperty> => {

    const diretoryPropertyName:string = profileCardProperties.directoryPropertyName;
    const _updateProfileCardProperty:IUpdateProfileCardProperty = { annotations : profileCardProperties.annotations};
    const _profileProperty: IProfileCardProperty = await msGraphClient
      .api(`/admin/people/profileCardProperties/${diretoryPropertyName}`)
      .patch(_updateProfileCardProperty);

      return _profileProperty;
  };

   // get Profile Card Property
   const getProfileCardProperty = async (
    msGraphClient: MSGraphClientV3,
    directoryPropertyName:string
  ):Promise<IProfileCardProperty> => {

    const _profileProperty: IProfileCardProperty = await msGraphClient
      .api(`/admin/people/profileCardProperties/${directoryPropertyName}`)
      .get();

      return _profileProperty;
  };

    // Delete Profile Card Property
    const deleteProfileCardProperty = async (
      msGraphClient: MSGraphClientV3,
      directoryPropertyName:string
    ) => {

      const _profileProperty: IProfileCardProperty = await msGraphClient
        .api(`/admin/people/profileCardProperties/${directoryPropertyName}`)
        .delete();
    };


    // check if user is Tenant Admin
    const checkUserIsGlobalAdmin  = async ( msGraphClient: MSGraphClientV3):Promise<boolean> =>  {
      const myDirRolesAndGroupsResults  =  await msGraphClient
      .api(`/me/memberof`)
      .get();
    const myDirRolesAndGroups = myDirRolesAndGroupsResults ? myDirRolesAndGroupsResults.value : [];
    for (const myDirRolesAndGroup of myDirRolesAndGroups) {
      if (myDirRolesAndGroup.roleTemplateId && myDirRolesAndGroup.roleTemplateId === ADMIN_ROLETEMPLATE_ID) { // roleTemplateId for glabal Admin
        return true;
      }
    }
    return false;
  };

  // return
  return { checkUserIsGlobalAdmin, getProfileCardProperties, newProfileCardProperty, updateProfileCardProperty, getProfileCardProperty, deleteProfileCardProperty };
};