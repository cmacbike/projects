using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Org.BouncyCastle.Bcpg;
using Org.BouncyCastle.Bcpg.OpenPgp;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Security;
namespace StarmineXML
{
    class EncryptSign
    {
        private string m_per_proc_dt;
        private string m_DBServerName;
        private string m_DatabaseName;
        private string m_Error;
        private bool m_IsError = false;
        private PgpPrivateKey privateKey;
        private PgpSecretKey secretKey;
        private PgpPublicKey encryptionKey;
        public void Encrypt(string inputFilePath, string recipientPublicKeyPath, string PrivateKeyPath, string passPhrase)
        {
            try
            {
                string l_inputFilePath = inputFilePath;
                //if (l_inputFilePath.Contains("YYYYMMDD"))
                //{
                //    SqlConnection connRIMS = getDBConnection(m_DBServerName, m_DatabaseName);



                //    SqlCommand cmdRIMS = new SqlCommand("select * from rims..proc_dt", connRIMS);
                //    SqlDataReader drRIMS = cmdRIMS.ExecuteReader();
                //    drRIMS.Read();
                //    DateTime dt_proc_dt = (DateTime)drRIMS[m_per_proc_dt];
                //    string proc_dt = String.Format("{0:yyyyMMdd}", dt_proc_dt);

                //    l_inputFilePath = l_inputFilePath.Replace("YYYYMMDD", proc_dt);
                //}
                //get data to encrypt 
                

                FileInfo infoInputFile = new FileInfo(l_inputFilePath);
                FileStream inputFile = infoInputFile.OpenRead();
                byte[] DataBuffer = new byte[infoInputFile.Length];


                //create memory stream to hold output from encryption  
                Stream finalOut = File.Create(l_inputFilePath + ".pgp");



                //get public key to encrypt message  
                encryptionKey = this.ReadPublicKey(recipientPublicKeyPath);


                //get secret key to sign message
                secretKey = this.ReadSecretKey(PrivateKeyPath);

                //get private key to sign message
                privateKey = this.ReadPrivateKey(passPhrase);




                //initialise encrypted data generator  
                PgpEncryptedDataGenerator encryptedDataGenerator = new PgpEncryptedDataGenerator(SymmetricKeyAlgorithmTag.Idea, new SecureRandom());
                encryptedDataGenerator.AddMethod(encryptionKey);
                Stream encOut = encryptedDataGenerator.Open(finalOut, DataBuffer.Length);


                //initialise compression  
                PgpCompressedDataGenerator compressedDataGenerator = new PgpCompressedDataGenerator(CompressionAlgorithmTag.Zip);
                Stream compressedOut = compressedDataGenerator.Open(encOut);



                //initialize Signature Generator
                PgpSignatureGenerator signatureGenerator = InitSignatureGenerator(compressedOut);

                PgpLiteralDataGenerator literalDataGenerator = new PgpLiteralDataGenerator();

                Stream literalOut = literalDataGenerator.Open(compressedOut, // the compressed output stream  
                                                        PgpLiteralData.Binary,
                                                        infoInputFile   // current time  
                                                        );

                //write data to output stream (eventually - goes through literal stream, compression stream, and encryption stream on the way!)  


                int length = 0;
                while ((length = inputFile.Read(DataBuffer, 0, DataBuffer.Length)) > 0)
                {
                    literalOut.Write(DataBuffer, 0, length);
                    signatureGenerator.Update(DataBuffer, 0, length);
                }
                signatureGenerator.Generate().Encode(compressedOut);


                //close literal output  
                literalOut.Close();
                literalDataGenerator.Close();



                //close other output streams  
                compressedOut.Close();
                compressedDataGenerator.Close();
                encOut.Close();
                encryptedDataGenerator.Close();
                finalOut.Close();
                inputFile.Close();
                //infoInputFile.Delete();
            }
            catch (Exception e)
            {
                m_IsError = true;
                m_Error = e.Message;
            }

        }

        private PgpSecretKey ReadSecretKey(string senderPrivateKeyPath)
        {

            Stream keyIn = File.OpenRead(senderPrivateKeyPath);
            Stream inputStream = PgpUtilities.GetDecoderStream(keyIn);
            PgpSecretKeyRingBundle pgpSec = new PgpSecretKeyRingBundle(inputStream);
            inputStream.Close();
            keyIn.Close();

            // just loop through the collection till we find a key suitable for encryption  
            // assuming only one key in there  

            foreach (PgpSecretKeyRing kRing in pgpSec.GetKeyRings())
            {
                foreach (PgpSecretKey k in kRing.GetSecretKeys())
                {
                    if (k.IsSigningKey)
                    {
                        return k;
                    }
                }
            }

            throw new ArgumentException("Can't find signing key in key ring.");
        }

        private PgpPublicKey ReadPublicKey(string publicKeyPath)
        {
            Stream keyIn = File.OpenRead(publicKeyPath);
            Stream inputStream = PgpUtilities.GetDecoderStream(keyIn);
            PgpPublicKeyRingBundle pgpPub = new PgpPublicKeyRingBundle(inputStream);
            inputStream.Close();
            keyIn.Close();

            foreach (PgpPublicKeyRing kRing in pgpPub.GetKeyRings())
            {
                foreach (PgpPublicKey k in kRing.GetPublicKeys())
                {
                    if (k.IsEncryptionKey)
                    {
                        return k;
                    }
                }
            }
            throw new ArgumentException("Can't find encryption key in key ring.");
        }
        private PgpPrivateKey ReadPrivateKey(string passPhrase)
        {


            PgpPrivateKey privateKey = secretKey.ExtractPrivateKey(passPhrase.ToCharArray());


            if (privateKey != null)


                return privateKey;


            throw new ArgumentException("No private key found in secret key.");


        }
        private PgpSignatureGenerator InitSignatureGenerator(Stream compressedOut)
        {


            const bool IsCritical = false;
            const bool IsNested = false;


            PublicKeyAlgorithmTag tag = secretKey.PublicKey.Algorithm;


            PgpSignatureGenerator pgpSignatureGenerator =  new PgpSignatureGenerator(tag, HashAlgorithmTag.Sha1);
            pgpSignatureGenerator.InitSign(PgpSignature.BinaryDocument, privateKey);

            foreach (string userId in secretKey.PublicKey.GetUserIds())
            {


                PgpSignatureSubpacketGenerator subPacketGenerator = new PgpSignatureSubpacketGenerator();
                subPacketGenerator.SetSignerUserId(IsCritical, userId);
                pgpSignatureGenerator.SetHashedSubpackets(subPacketGenerator.Generate());


                // Just the first one!


                break;


            }
            pgpSignatureGenerator.GenerateOnePassVersion(IsNested).Encode(compressedOut);
            return pgpSignatureGenerator;
        }


    

        //private SqlConnection getDBConnection(string p_DBServerName, string p_DatabaseName)
        //{
        //    string strConnectionString = String.Format("Server={0};Database={1};Trusted_Connection=True", p_DBServerName, p_DatabaseName);
        //    SqlConnection connRIMS = new SqlConnection(strConnectionString);
        //    connRIMS.Open();
        //    return connRIMS;

        //}
        public string per_proc_dt
        {
            get
            {
                return m_per_proc_dt;
            }
            set
            {
                m_per_proc_dt = value;

            }
        }
        public string DBServerName
        {
            get
            {
                return m_DBServerName;
            }
            set
            {
                m_DBServerName = value;

            }
        }
        public string DatabaseName
        {
            get
            {
                return m_DatabaseName;
            }
            set
            {
                m_DatabaseName = value;

            }
        }
        public string Error
        {
            get
            {
                return m_Error;
            }
            set
            {
                m_Error = value;

            }
        }
        public bool IsError
        {
            get
            {
                return m_IsError;
            }
            set
            {
                m_IsError = value;

            }
        }
    }
}
