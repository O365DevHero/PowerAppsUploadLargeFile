import { BlobServiceClient, HttpHeaders } from "@azure/storage-blob";
import { AbortController } from "@azure/abort-controller";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as fs from "fs";

export class CanvasFileUploader
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _container: HTMLDivElement;
  private _fileInput: HTMLInputElement;
  private _notifyOutputChanged: () => void;
  private _outputType: string;
  private _value: string;
  private _fileName: string;

  private _blockIds: string[];
  private _counter: number;

  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ) {
    this._notifyOutputChanged = notifyOutputChanged;

    this._value = "";
    this._fileName = "";
    this._outputType = context.parameters.outputType.raw || "DataUrl";

    this._fileInput = document.createElement("input");
    this._fileInput.type = "file";
    this._fileInput.style.display = context.mode.isVisible ? "block" : "none";

    this._fileInput.onchange = this.onFileInputChange.bind(this);
    container.appendChild(this._fileInput);
  }

  public onFileInputChange = () => {
    if (this._fileInput.files?.length === 0) {
      this._notifyOutputChanged();
      return;
    }

    let file = this._fileInput.files![0];
    this._fileName = file.name;
    this._blockIds = [];
    this._counter = 1;
    // this.parseFile(file, () => {
    //   console.log("In callback");
    // });

    this.uploadBlobsInChunks(file);

    // var fileReader = new FileReader();
    // fileReader.onloadend = () => {
    // 	this._value = fileReader.result as string;
    // 	console.log(this._value);

    // 	this.uploadToBlob();
    // 	this._notifyOutputChanged();
    // }

    // if (this._outputType === 'Text')
    // {
    // 	fileReader.readAsText(file);
    // }
    // else{
    // 	fileReader.readAsDataURL(file);
    // }
  };

  public uploadBlobsInChunks = async (file: File) => {
    let fileSize = file.size;
    let chunkSize = 4 * 1024 * 1024; // bytes
    let offset = 0;
    const account = "sciicqw";
    const sas =
      "?sv=2020-08-04&ss=bfqt&srt=sco&sp=rwdlacuptfx&se=2021-09-02T20:47:44Z&st=2021-08-27T12:47:44Z&spr=https&sig=sz3JyZba9dGQhsgy4mZxydp2RKvJksuTQ6ImFOlqzc4%3D";
    const containerName = "sciisubmissions";
    const blobServiceClient = new BlobServiceClient(
      `https://${account}.blob.core.windows.net${sas}`
    );
    var containerClient = blobServiceClient.getContainerClient(containerName);
    var blockBlobClient = containerClient.getBlockBlobClient(this._fileName);
    await blockBlobClient.uploadData(file, {
      blockSize: 4 * 1024 * 1024, // 4MB block size
      concurrency: 5, // 20 concurrency
      onProgress: (ev) => console.log(ev),
    });
  };

  //Uploads to blob in single go
  public uploadToBlob = async () => {
    //const connStr = "DefaultEndpointsProtocol=https;AccountName=sciicqw;AccountKey=mdMDcK/DHEIHBNbo0ycIeiyejvQlR8om0WKWlNUFXiocZDf9Y0kjlpNHSJllK/Ftc9QeX1NuetnGOaYCpEk+Zw==;EndpointSuffix=core.windows.net;";
    const account = "sciicqw";
    const sas =
      "?sv=2020-08-04&ss=bfqt&srt=sco&sp=rwdlacuptfx&se=2021-08-27T09:12:24Z&st=2021-08-26T01:12:24Z&spr=https&sig=ZRKBqyBYs2Mm%2BJr5IOeow71kLtldbBvdjWZEXU9yM4U%3D";
    const containerName = "sciisubmissions";
    const blobServiceClient = new BlobServiceClient(
      `https://${account}.blob.core.windows.net${sas}`
    );

    //const blobServiceClient = BlobServiceClient.fromConnectionString(connStr);
    const containerClient = blobServiceClient.getContainerClient(containerName);
    const blobName = "newblob_" + new Date().getTime() + this._fileName;
    const blockBlobClient = containerClient.getBlockBlobClient(blobName);
    const uploadBlobResponse = await blockBlobClient.upload(
      this._value,
      this._value.length
    );
    console.log(
      `Upload block blob ${blobName} successfully.`,
      uploadBlobResponse.requestId
    );
  };

  //Upload to Blob in chunks and commit at the end
  public parseFile = (file: File, callback: Function) => {
    let fileSize = file.size;
    let chunkSize = 4 * 1024 * 1024; // bytes
    let offset = 0;

    let self = this; // we need a reference to the current object
    //let chunkReaderBlock = null;

    // now let's start the read with the first block
    this.chunkReaderBlock(offset, chunkSize, file);
  };

  public chunkReaderBlock = (_offset: number, length: number, _file: File) => {
    var r = new FileReader();
    var blob = _file.slice(_offset, length + _offset);
    const account = "sciicqw";
    const sas =
      "?sv=2020-08-04&ss=bfqt&srt=sco&sp=rwdlacuptfx&se=2021-08-27T09:12:24Z&st=2021-08-26T01:12:24Z&spr=https&sig=ZRKBqyBYs2Mm%2BJr5IOeow71kLtldbBvdjWZEXU9yM4U%3D";
    const containerName = "sciisubmissions";
    const blobServiceClient = new BlobServiceClient(
      `https://${account}.blob.core.windows.net${sas}`
    );

    r.onload = async (evt: any) => {
      if (evt.target.error == null) {
        var containerClient =
          blobServiceClient.getContainerClient(containerName);
        var blockBlobClient = containerClient.getBlockBlobClient(
          this._fileName
        );
        var blockId = window.btoa(
          encodeURIComponent(
            this._counter.toLocaleString("en-us", {
              minimumIntegerDigits: 6,
              useGrouping: false,
            })
          )
        );
        console.log(this._counter + "-->" + blockId);
        try {
          await blockBlobClient.stageBlock(blockId, blob, length, {
            abortSignal: AbortController.timeout(30 * 60 * 1000), // Abort uploading with timeout in 30mins
            onProgress: (ev) => console.log(ev),
          });

          this._counter = this._counter + 1;
          _offset += evt.target.result.length;
          this._blockIds.push(blockId);
        } catch (err) {
          console.log(
            `uploadStream failed, requestId - ${err.details.requestId}, statusCode - ${err.statusCode}, errorCode - ${err.details.errorCode}`
          );
        }
        //callback(evt.target.result); // callback for handling read chunk
      } else {
        console.log("Read error: " + evt.target.error);
        return;
      }
      if (_offset >= _file.size) {
        console.log("Done reading file.");
        blockBlobClient.commitBlockList(this._blockIds, {
          blobHTTPHeaders: { blobContentType: _file.type },
        });
        return;
      }
      this.chunkReaderBlock(_offset, length, _file);
    };
    r.readAsText(blob);
  };

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    var params = context.parameters;
    this._fileInput.accept = params.acceptedFileTypes.raw || "";
    this._fileInput.style.display = context.mode.isVisible ? "block" : "none";
    this._outputType = context.parameters.outputType.raw || "DataUrl";
    if (!params.triggerFileSelector?.raw) return;
    this._fileInput.click();
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return {
      value: this._value,
      fileName: this._fileName,
      triggerFileSelector: false,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }
}
