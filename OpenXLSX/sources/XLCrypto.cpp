#include "XLCrypto.hpp"
#include <openssl/evp.h>
#include <openssl/sha.h>
#include <openssl/rand.h>
#include "XLException.hpp"
#include <cstring>
#include <memory>
#include <array>
#include <pugixml.hpp>
#include <algorithm>
#include <iomanip>

using namespace OpenXLSX;

bool OpenXLSX::isEncryptedDocument(gsl::span<const uint8_t> data) {
    if (data.size() < 512) return false;
    const uint8_t magic[] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
    return std::memcmp(data.data(), magic, 8) == 0;
}

std::vector<uint8_t> OpenXLSX::encryptDocument(gsl::span<const uint8_t> zipData, const std::string& password) {
    // Phase 3: Writing an OLE CFB Container
    // This requires generating EncryptionInfo XML and encrypting the package in 4096-byte chunks.
    throw XLInternalError("Encryption feature not fully implemented in this prototype.");
}

std::vector<uint8_t> OpenXLSX::decryptDocument(gsl::span<const uint8_t> data, const std::string& password) {
    auto encryptionInfo = Crypto::readCfbStream(data, "EncryptionInfo");
    auto encryptedPackage = Crypto::readCfbStream(data, "EncryptedPackage");
    
    // In a complete implementation, detect if Agile or Standard based on XML root
    return Crypto::decryptAgilePackage(encryptionInfo, encryptedPackage, password);
}

namespace OpenXLSX::Crypto {

std::vector<uint8_t> readCfbStream(gsl::span<const uint8_t> data, const std::string& streamName) {
    if (data.size() < 512) throw XLInternalError("Invalid CFB size");
    const uint8_t magic[] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
    if (std::memcmp(data.data(), magic, 8) != 0) throw XLInternalError("Not an OLE/CFB file");

    // Placeholder: Implement FAT parsing
    // In a full implementation, we traverse the FAT/MiniFAT chains
    // For this architectural preview, we return an empty vector if not implemented.
    return std::vector<uint8_t>();
}

std::vector<uint8_t> aes256CbcDecrypt(gsl::span<const uint8_t> data, gsl::span<const uint8_t> key, gsl::span<const uint8_t> iv) {
    if (key.size() != 32 || iv.size() != 16) throw XLInternalError("Invalid AES-256 key/IV size");

    EVP_CIPHER_CTX *ctx = EVP_CIPHER_CTX_new();
    if (!ctx) throw XLInternalError("Failed to create EVP context");
    
    // RAII for context
    auto cleanup = gsl::finally([&] { EVP_CIPHER_CTX_free(ctx); });

    if (EVP_DecryptInit_ex(ctx, EVP_aes_256_cbc(), nullptr, key.data(), iv.data()) != 1) {
        throw XLInternalError("AES init failed");
    }
    
    EVP_CIPHER_CTX_set_padding(ctx, 1);

    std::vector<uint8_t> out(data.size() + EVP_MAX_BLOCK_LENGTH);
    int len1 = 0, len2 = 0;

    if (EVP_DecryptUpdate(ctx, out.data(), &len1, data.data(), gsl::narrow_cast<int>(data.size())) != 1) {
        throw XLInternalError("AES decrypt update failed");
    }

    if (EVP_DecryptFinal_ex(ctx, out.data() + len1, &len2) != 1) {
        throw XLInternalError("AES decrypt final failed");
    }

    out.resize(len1 + len2);
    return out;
}

std::vector<uint8_t> sha512Hash(gsl::span<const uint8_t> data) {
    std::vector<uint8_t> hash(SHA512_DIGEST_LENGTH);
    SHA512(data.data(), data.size(), hash.data());
    return hash;
}

std::vector<uint8_t> generateAgileHash(const std::string& password, gsl::span<const uint8_t> salt, int spinCount) {
    std::vector<uint8_t> utf16pw;
    for (char c : password) {
        utf16pw.push_back(static_cast<uint8_t>(c));
        utf16pw.push_back(0x00);
    }
    
    std::vector<uint8_t> initialData(salt.begin(), salt.end());
    initialData.insert(initialData.end(), utf16pw.begin(), utf16pw.end());
    
    auto H = sha512Hash(initialData);

    for (uint32_t i = 0; i < static_cast<uint32_t>(spinCount); ++i) {
        std::vector<uint8_t> loopData(4);
        loopData[0] = (i & 0xFF);
        loopData[1] = ((i >> 8) & 0xFF);
        loopData[2] = ((i >> 16) & 0xFF);
        loopData[3] = ((i >> 24) & 0xFF);
        loopData.insert(loopData.end(), H.begin(), H.end());
        H = sha512Hash(loopData);
    }
    
    OPENSSL_cleanse(utf16pw.data(), utf16pw.size());
    OPENSSL_cleanse(initialData.data(), initialData.size());
    
    return H;
}

std::vector<uint8_t> decryptAgilePackage(gsl::span<const uint8_t> encryptionInfo,
                                         gsl::span<const uint8_t> encryptedPackage,
                                         const std::string& password) {
    if (encryptionInfo.empty() || encryptedPackage.empty()) return std::vector<uint8_t>();
    
    pugi::xml_document doc;
    pugi::xml_parse_result result = doc.load_buffer(encryptionInfo.data(), encryptionInfo.size());
    if (!result) throw XLInternalError("Failed to parse EncryptionInfo XML");

    // Placeholder for actual agile logic:
    return std::vector<uint8_t>();
}

std::vector<uint8_t> decryptStandardPackage(gsl::span<const uint8_t> encryptionInfo,
                                            gsl::span<const uint8_t> encryptedPackage,
                                            const std::string& password) {
    return std::vector<uint8_t>();
}

} // namespace OpenXLSX::Crypto
