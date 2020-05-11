/*
    <copyright file="nominate-awards-api.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import axios from "./axios-decorator";
const baseAxiosUrl = window.location.origin;

/**
* Save nominated details.
* @param  {NominateEntity | Null} nominateDetails nominated details.
*/
export const saveNominateDetails = async (nominateDetails: any): Promise<any> => {

    let url = baseAxiosUrl + "/api/NominateDetail/nomination";
    return await axios.post(url, nominateDetails, undefined);
}

/**
* Get nominated award details.
* @param  {String | Null} teamId Team id.
* @param  {String | Null} aadObjectId User azure active directory object id.
* @param  {String | Null} cycleId Active award cycle unique id.
*/
export const getNominationAwardDetails = async (teamId: string | null, aadObjectId: string | null, cycleId: string | null): Promise<any> => {
    let url = baseAxiosUrl + `/api/NominateDetail/nominationdetail?teamId=${teamId}&aadObjectId=${aadObjectId}&cycleId=${cycleId}`;
    return await axios.get(url, undefined);
}

/**
* Get all nominations from API.
* @param {String} teamId Team Id for which the awards will be fetched.
 *@param {boolean} isAwardGranted flag: true for published award, else false.
 *@param {String} awardCycleId Active award cycle unique id.
*/
export const getAllAwardNominations = async (teamId: string | undefined, isAwardGranted: boolean | undefined, awardCycleId: string | undefined): Promise<any> => {

    let url = baseAxiosUrl + `/api/NominateDetail/allnominations?teamId=${teamId}&isAwardGranted=${isAwardGranted}&awardCycleId=${awardCycleId}`;
    return await axios.get(url, undefined);
}

/**
* publish nominations from API
* @param {String} teamId Team Id for which the awards will be fetched.
 *@param {String} nominationIds Publish nomination ids.
*/
export const publishAwardNominations = async (teamId: string | undefined, nominationIds: string | undefined): Promise<any> => {

    let url = baseAxiosUrl + `/api/NominateDetail/publishnominations?teamId=${teamId}&nominationIds=${nominationIds}`;
    return await axios.get(url, undefined);
}