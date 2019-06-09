import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as Jodit from "./src/jodit.min.js";

export class JoditRichTextController implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Value of the field is stored and used inside the control 
    private _value: string;
    // PowerApps component framework framework delegate which will be assigned to this object which would be called whenever an update happens. 
    private _notifyOutputChanged: () => void;
    // input element that is used to create the rich text controller
    private inputElement: HTMLTextAreaElement;
    // Reference to the control container HTMLDivElement
    private _container: HTMLDivElement;
    // Reference to ComponentFramework Context object
    private _context: ComponentFramework.Context<IInputs>;
    // Event Handelr 'refreshData' reference
    private _refreshJoditData: EventListenerOrEventListenerObject;
	// Jodit richtext controller
	private _richTextEditor: Jodit;
	// Jodit richtect controller's default value
	private _default: string;
	//Class and ID names
	private _className: string;


	/**
	 * JoditRichTextController constructor.
	 */
	constructor()
	{
		this._default = '<p style="text-align: left;"><br></p>';
		this._className = 'jodit_richtextinput';
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Control initialization code
        this._context = context;
		this._container = document.createElement("div");
		this._container.setAttribute("class", "joditPCFContainer");
        this._notifyOutputChanged = notifyOutputChanged;
		this._refreshJoditData = this.refreshJoditData.bind(this);
		
		// creating HTML elements for the input type range and binding it to the function which refreshes the control data
        this.inputElement = document.createElement("textarea");
		
		let randomInt: number = Math.floor(Math.floor(100) * Math.random());
		this._className = this._className + randomInt;
        //setting the display class and id for the control.
        this.inputElement.setAttribute("class", this._className);
        this.inputElement.setAttribute("id", this._className);
        
        // retrieving the latest value from the control and setting it to the HTMl elements.
        this._value = context.parameters.JoditRichTextProperty.raw;

        // appending the HTML elements to the control's HTML container element.
		this._container.appendChild(this.inputElement);
		
        container.appendChild(this._container);

        //initiate Jodit Editor to the textarea
        this._richTextEditor = new Jodit('#'+ this._className, {
			minHeight: '350'
		 });
        //add event listener for on change event
        this._richTextEditor.events.on('change',this._refreshJoditData);
        //Setting the editor default value
        this._richTextEditor.setEditorValue(context.parameters.JoditRichTextProperty.raw ? context.parameters.JoditRichTextProperty.raw : this._default);
        
	}

	/**
	 * Updates the values to the internal value variable we are storing
	 * @param evt : The "Input Properties" containing the parameters, control metadata and interface functions
	 */
	public refreshJoditData(evt: Event): void
    {
		if(this._value != this._richTextEditor.value){
			this._value = this._richTextEditor.value;
			this._notifyOutputChanged();
		}        
    }

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Storing the latest value and context from the control.
		if(this._value != context.parameters.JoditRichTextProperty.raw){
			this._value = context.parameters.JoditRichTextProperty.raw;
			//this._richTextEditor.setEditorValue(context.parameters.JoditRichTextProperty.raw ? context.parameters.JoditRichTextProperty.raw : this._default);
		}		
		this._context = context;
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
            JoditRichTextProperty: this._value
        };
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// cleanup control
        this.inputElement.removeEventListener("change", this._refreshJoditData);
        this._richTextEditor.destroy();
	}
}