const fs = require('fs')
const path = require('path')
const dotenv = require('dotenv')
const fetch = require('node-fetch')
const DomParser = require('dom-parser');
const { decode } = require("html-entities")

//===========[ Configuration ]===========
const SCHOOL_ID = "" // can be configured using environment variable 'SCHOOL'

const downloadLocation = "./data"
const skipAlreadyDownloadedFiles = true // skip files, if they already exists
//=======================================

async function main() {
    dotenv.config()

    const school = process.env.SCHOOL ? process.env.SCHOOL : SCHOOL_ID
    if (!school) {
        console.error("School not specified")
        process.exit(1)
    }

    let sessionId = process.env.SESSION

    if (!sessionId) {
        const username = process.env.USERNAME
        if (!username) {
            console.error("Username not specified")
            process.exit(1)
        }

        const password = process.env.PASSWORD
        if (!password) {
            console.error("Password not specified")
            process.exit(1)
        }

        sessionId = await loginForSessionId(school, username, password)
    }

    const courses = await getCourses(school, sessionId)

    const courseNames = courses.map(c => c.Title)
    console.log("Available courses: ")
    console.log(courseNames)

    await downloadResourcesOfCourses(school, sessionId, courses)
}

async function loginForSessionId(school, username, password) {
    const response = await fetch(`https://${school}.itslearning.com`, {
        method: 'post',
        body: `ctl00$ContentPlaceHolder1$Username$input=${username}&ctl00$ContentPlaceHolder1$Password$input=${password}`
    })

    const cookies = response.headers.raw()['set-cookie'].filter(c => c.includes("ASP.NET_SessionId"))
    const sessionId = cookies[0].match(/ASP\.NET_SessionId=(\S+);/)[1]

    console.log("Obtained session id: " + sessionId)

    if (!sessionId) {
        console.error("Failed to obtain session id")
        process.exit(1)
    }
    return sessionId
}

async function getCourses(school, sessionId) {
    const baseUrl = `https://${school}.itslearning.com/restapi/personal/courses/cards/`
    const starredUrl = baseUrl + "starred/v1?SortBy=LastUpdated"
    const unstarredUrl = baseUrl + "unstarred/v1?SortBy=LastUpdated"

    const options = {
        method: 'get',
        headers: {
            'Content-Type': 'application/json',
            'Cookie': `ASP.NET_SessionId=${sessionId};`
        }
    }

    const responseStarred = await fetch(starredUrl, options)
    if (!responseStarred.ok) throw new Error("Cannot get starred courses: " + responseStarred.statusText)
    const starredCourses = await responseStarred.json()

    const responseUnstarred = await fetch(unstarredUrl, options)
    if (!responseUnstarred.ok) throw new Error("Cannot get unstarred courses: " + responseUnstarred.statusText)
    const unstarredCourses = await responseUnstarred.json()

    return starredCourses.EntityArray.concat(unstarredCourses.EntityArray)
}

async function downloadResourcesOfCourses(school, sessionId, courses) {
    for (const course of courses) {
        const contentUrl = `https://${school}.itslearning.com/ContentArea/ContentArea.aspx?LocationID=${course.CourseId}&LocationType=1`

        const contentResponse = await fetch(contentUrl, {
            headers: {
                'Cookie': `ASP.NET_SessionId=${sessionId};`
            }
        })

        if (!contentResponse.ok) {
            console.error("Failed to fetch web content for course " + course.Title)
            return
        }

        const webContent = await contentResponse.text()

        const parser = new DomParser()
        const dom = parser.parseFromString(webContent)

        const resourcesAnchor = dom.getElementById("link-resources")
        const resourceFolderUrl = resourcesAnchor.getAttribute("href")

        const parentFolder = path.join(downloadLocation, decodeString(course.Title))
        await downloadResourceFolder(school, resourceFolderUrl, sessionId, parentFolder)
    }
}

/**
 * Downloads all files in the given folder recursively.
 * @param {String} folderUrl url refering to the target folder
 * @param {String} sessionId session ID for <school>.itslearning.com
 * @param {String} parentFolderPath the relative path of the parent folder
 * @param {Boolean} appendId set to true to append the folder ID to the name
 *                           if there are multiple folders with the same name
 */
async function downloadResourceFolder(school, folderUrl, sessionId, parentFolderPath, appendId) {
    if (folderUrl.startsWith("/")) {
        folderUrl = `https://${school}.itslearning.com${folderUrl}`
    }

    const contentResponse = await fetch(folderUrl, {
        headers: {
            'Cookie': `ASP.NET_SessionId=${sessionId};`
        }
    })

    if (!contentResponse.ok) {
        console.error("Failed to fetch web content for folder at " + folderUrl)
        return
    }

    const webContent = await contentResponse.text()

    const parser = new DomParser()
    const dom = parser.parseFromString(webContent)

    // extract current folder name
    const rawFolderName = dom.getElementById("ctl00_PageHeader_TT").textContent
    let folderName = decodeString(rawFolderName)
    if (appendId) {
        // since multiple folders can have the same name, we append the folder ID
        const folderId = folderUrl.match(/FolderID=(\d+)/)[1]
        folderName = `${folderName} [${folderId}]`
    }
    const folderPath = `${parentFolderPath}/${folderName}`

    const folderEntries = dom.getElementsByClassName("GridTitle")
    const folderEntryNames = folderEntries.map(e => e.textContent)

    for (const entry of folderEntries) {
        const url = entry.getAttribute("href")

        const numberOfSameName = folderEntryNames.filter(n => n === entry.textContent).length
        const shouldAppendId = numberOfSameName > 1 ? true : false

        if (url.startsWith("/Folder")) {
            // go to nested folder
            await downloadResourceFolder(school, url, sessionId, folderPath, shouldAppendId)
        } else if (url.startsWith("/LearningToolElement")) {
            // extract element id and download file
            const elementId = url.match(/LearningToolElementId=(\d+)/)[1]
            downloadResourceObject(school, sessionId, elementId, folderPath)
        } else {
            console.warn("Unknown folder entry type. Url: " + url)
        }
    }
}

/**
 * Extracts the download url for the resource of the given element ID and starts the actual file download.
 * @param {String} school school ID
 * @param {String} sessionId session ID for <school>.itslearning.com
 * @param {String|Number} elementId resource object ID
 * @param {String} folderPath parent folder path
 * @returns 
 */
async function downloadResourceObject(school, sessionId, elementId, folderPath) {
    const url = `https://${school}.itslearning.com/LearningToolElement/ViewLearningToolElement.aspx?LearningToolElementId=${elementId}`

    const viewResponse = await fetch(url, {
        method: 'get',
        redirect: 'manual',
        headers: {
            'Cookie': `ASP.NET_SessionId=${sessionId};`
        }
    })

    if (!viewResponse.ok) {
        console.error(`Failed to fetch web page for resource object: ${viewResponse.statusText} URL: ${url}`)
        return
    }

    const viewWebContent = await viewResponse.text()
    const parser = new DomParser()
    const viewDom = parser.parseFromString(viewWebContent)

    const fileName = viewDom.getElementById("ctl00_PageHeader_TT").textContent

    const iframe = viewDom.getElementById("ctl00_ContentPlaceHolder_ExtensionIframe")
    const iframeUrl = iframe.getAttribute("src").replace(/&amp;/g, "&")

    //console.log("iframe url: " + iframeUrl)

    // this will obtain the session cookie for platform.itslearning.com
    const iframeResponse = await fetch(iframeUrl, {
        method: 'get',
        redirect: 'manual'
    })
    const learningObjectInstanceUrl = iframeResponse.headers.get('location')

    // extract platform cookies for following requests
    const platformCookieString = convertSetCookieToCookieString(iframeResponse.headers.raw()['set-cookie'])

    //console.log("learning object instance url: " + learningObjectInstanceUrl)

    // this should redirect us to the resource page
    const resourceResponse = await fetch(learningObjectInstanceUrl, {
        method: 'get',
        headers: {
            'Cookie': platformCookieString
        }
    })

    // extract cookies
    const resourceCookies = resourceResponse.headers.raw()['set-cookie'].filter(c => c.includes("ASP.NET_SessionId"))
    const resourceCookieString = convertSetCookieToCookieString(resourceCookies)

    const resourceWebContent = await resourceResponse.text()
    const resourceDom = parser.parseFromString(resourceWebContent)

    // direct download link for e.g. PDF files
    const downloadAnchor = resourceDom.getElementById("ctl00_ctl00_MainFormContent_DownloadLinkForViewType")
    // Office for web iframe
    const filePreviewIframe = resourceDom.getElementById("ctl00_ctl00_MainFormContent_PreviewIframe_FilePreviewIframe")

    if (downloadAnchor) {
        // DIRECT DOWNLOAD
        const downloadUrl = "https://resource.itslearning.com" + downloadAnchor.getAttribute("href").replace(/&amp;/g, "&")
        const downloadFilename = downloadAnchor.getAttribute("Download")

        //console.log("download url: " + downloadUrl)
        await downloadFile(downloadUrl, resourceCookieString, folderPath, downloadFilename)
    } else if (filePreviewIframe) {
        // OFFICE FILE DOWNLOAD
        const previewUrl = "https://resource.itslearning.com/" + filePreviewIframe.getAttribute("src")

        const previewResponse = await fetch(previewUrl, {
            headers: {
                'Cookie': resourceCookieString
            }
        })

        if (!previewResponse.ok) {
            console.error("Failed to download file preview web page (office online): " + previewResponse.statusText)
            return
        }

        const previewContent = await previewResponse.text()
        const officeDownloadUrl = getOfficeFileDownloadUrl(previewContent)

        await downloadFile(officeDownloadUrl, "", folderPath, fileName)
    } else {
        console.log("Resource type unsupported. No direct download link available: " + fileName)
        return
    }
}

function getOfficeFileDownloadUrl(webContent) {
    const officeUrl = webContent.match(/form\.action = '(\S+)';/)[1]
    //console.log("Office url: " + officeUrl);

    const accessToken = webContent.match(/accessTokenInput\.value = '(\S+)';/)[1]
    const accessTokenTtl = webContent.match(/accessTokenTtlInput\.value = '(\S+)';/)[1]

    //console.log("access token: " + accessToken);
    //console.log("access token ttl: " + accessTokenTtl);

    const fileContentUrl = officeUrl.match(/WOPISrc=(\S+)&ui=/)[1].replace(/\\x253a/g, ":").replace(/\\x252f/g, "/")
    const fileDownloadUrl = `${fileContentUrl}/contents?access_token=${accessToken}`
    //console.log("download url: " + fileDownloadUrl);

    return fileDownloadUrl

    // Fetch office site and extract download URL from 'DownloadACopyUrl' attribute
    // const response = await fetch(officeUrl, {
    //     method: 'post',
    //     body: `access_token=${accessToken}&access_token_ttl=${accessTokenTtl}`
    // })
    // const officeContent = await response.text()
    // const downloadCopyUrl = officeContent.match(/DownloadACopyUrl: '(\S+)'/)[1]
}

/**
 * Downloads the file to the local machine.
 * @param {String} downloadUrl the url refering to the resource
 * @param {String} cookieString cookie containing the session ID for resource.itslearning.com
 * @param {String} folderPath parent folder path for the file
 * @param {String} filename the target file name
 */
async function downloadFile(downloadUrl, cookieString, folderPath, filename) {
    const downloadLocation = path.join(folderPath, filename)

    if (fs.existsSync(downloadLocation) && skipAlreadyDownloadedFiles) {
        console.log("Skipping file: " + downloadLocation)
        return
    }

    const response = await fetch(downloadUrl, {
        headers: {
            'Cookie': cookieString
        }
    })

    if (!response.ok) {
        console.log("Failed to download file: " + response.statusText)
        return
    }

    if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath, { recursive: true })
    }

    const fileStream = fs.createWriteStream(downloadLocation)

    await new Promise((resolve, reject) => {
        response.body.pipe(fileStream)
        response.body.on("error", reject)
        fileStream.on("finish", resolve)
    })

    console.log("Downloaded file to " + downloadLocation)
}

function convertSetCookieToCookieString(cookieArray) {
    const cookies = cookieArray.map(cookie => cookie.substring(0, cookie.indexOf(";")))
    const uniqueCookies = [...new Set(cookies)]
    return uniqueCookies.join('; ')
}

/**
 * Decode text as UTF-8 and replace special characters that are not
 * allowed in file names.
 */
function decodeString(text) {
    const decoded = decode(Buffer.from(text, 'utf-8').toString())
    return decoded.replace(/[\/\\|":?*<>{}]/g, '_')
}

main()
