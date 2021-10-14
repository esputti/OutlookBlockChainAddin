/* global Office */
import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import $ from "jquery";
import Web3 from "../web3.min.js";
import { Account_Email_List } from "./constants.js"

/* global require */

var web3 = new Web3(new Web3.providers.HttpProvider("http://127.0.0.1:7545"));

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      web3Instance: null,
      web3Provider: null,
      contracts: {},
      account: "0x0",
      loading: false,
      contractInstance: null,
      msg: "0x0",
      signature: "0x0",
      accountmessage: "",
      signatureVerificationMessage: "",
    };

    this.signMessage = this.signMessage.bind(this);
    this.signMessageCallback = this.signMessageCallback.bind(this);
    this.verifySignature = this.verifySignature.bind(this);
    this.verifySignatureCallback = this.verifySignatureCallback.bind(this);
    this.checkContractInstance = this.checkContractInstance.bind(this);
  }

  componentDidMount() {
    this.setState({
      listItems: [],
    });
    this.initWeb3();
  }

  signMessage = async () => {
    console.log('In Sign message');
    await this.initWeb3();
    Office.context.mailbox.item.body.getAsync("text", {}, this.signMessageCallback);
  };

  signMessageCallback = async (result) => {
    // Do something with the result.
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("error in getting message body");
    }
    var item = result.value.trim();
    console.log("item : ", item);
    console.log("item body: ", item);

    const email = Office.context.mailbox.userProfile.emailAddress;
    var account = "";
    for (var i = 0; i < Account_Email_List.length; i++) {
      if (Account_Email_List[i].Email === email) {
        account = Account_Email_List[i].Account;
      }
    }
    const message = web3.utils.sha3(item);
    console.log('message', message);
    this.setState({ msg: message });


    console.log("account before signing: ", account);
    this.setState({ accountmessage: "Your account: " + account + ", " + email });

    let sig1 = await web3.eth.sign(message, account);

    console.log("Signature: ", sig1);
    this.setState({ signature: sig1 });
    Office.context.mailbox.item.body.appendOnSendAsync(
      "DigitalSignature:" + sig1,
      {coercionType: Office.CoercionType.Text},
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log("Appendonsendasync failed with error: " + asyncResult.error.message);
        }
      }
    );
  };

  verifySignature = async () => {
    await this.initWeb3();
    this.checkContractInstance();
  };

  checkContractInstance = () => {
    if (this.state.contractInstance != null) {
      console.log("emailAddress: ", Office.context.mailbox.userProfile.emailAddress);
      console.log("In verify signature");
      console.log("contract instance: ", this.state.contractInstance);
      Office.context.mailbox.item.body.getAsync("text", {}, this.verifySignatureCallback);
    } else {
      window.setTimeout(this.checkContractInstance, 100);
    }
  };

  verifySignatureCallback = async (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("error in getting message body");
    }
    var signedMessage = result.value;
    console.log("item body: ", signedMessage);
    var index = signedMessage.lastIndexOf("DigitalSignature:");
    if (index == -1) {
      this.setState({ signatureVerificationMessage: "Email is not digitally signed." });
    }
    const signature = signedMessage.substr(index + 17).trim();
    const msg = signedMessage.substr(0, index).trim();
    console.log("item body with out signature: ", msg);
    console.log("signature from item body: ", signature);
    const shamessage = web3.utils.sha3(msg);
    console.log('shamessage: ', shamessage);
    console.log("verify signature callback contractInstance: ", this.state.contractInstance);

    this.state.contractInstance.recover(shamessage, signature)
      .then((result) => {
        console.log('Recover: ', result);
        var email = "";
        for (var i = 0; i < Account_Email_List.length; i++){
          if (Account_Email_List[i].Account === result) {
            email = Account_Email_List[i].Email;
          }
        }
        this.setState({ 
          signatureVerificationMessage: "This email is signed by " + result + ", " + email,
        });
      })
      .catch((err) => { console.log("There was an error recovering signature.", err); });
  };

  initWeb3 = async () => {
    this.setState({
      web3Provider: web3.currentProvider,
    });
    console.log("web3provider: ", web3.currentProvider);
    console.log("In initweb3");
    var verificationcontract = require("../Verification.json");
    console.log("contract", verificationcontract);
    this.state.contracts.Verification = TruffleContract(verificationcontract);
    this.state.contracts.Verification.setProvider(web3.currentProvider);

    // var accounts = await web3.eth.getAccounts();
    // console.log("All accounts: ", accounts);
    // var acc = await accounts[0];
    // this.setState({ account: acc });
    // console.log("Your Account:", this.state.account);

    const email = Office.context.mailbox.userProfile.emailAddress;
    console.log("In initweb3, email : ", email);
    var acc = "";
    for (var i = 0; i < Account_Email_List.length; i++) {
      if (Account_Email_List[i].Email == email) {
        acc = Account_Email_List[i].Account;
      }
    }
    this.setState({ account: acc });
    this.setState({ web3Instance: web3, accountmessage: "Your account: " + acc + ", " + email });

    this.state.contracts.Verification.deployed()
      .then((contract) => {
        this.setState({ contractInstance: contract });
        console.log("ContractInstance", contract);
        console.log("Contract Address:", contract.address);
        return true;
      })
      .then((val) => {});
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">{this.state.accountmessage}</p>
          <p className="ms-font-l">{this.state.signatureVerificationMessage}</p>
          <DefaultButton className="ms-welcome__action" onClick={this.signMessage}>
            Sign Message
          </DefaultButton>
          <span> </span>
          <DefaultButton className="ms-welcome__action" onClick={this.verifySignature}>
            Verify Message
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
