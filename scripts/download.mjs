import * as stream from 'stream';
import * as fs from 'fs';
import axios, * as others from 'axios';
import { promisify } from 'util';

const finished = promisify(stream.finished);

export async function downloadFile(fileUrl, outputLocationPath) {
  const writer = fs.createWriteStream(outputLocationPath);
  return axios({
    method: 'get',
    url: fileUrl,
    responseType: 'stream',
  }).then(response => {
    response.data.pipe(writer);
    return finished(writer);
  });
}