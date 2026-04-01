#include <map>
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
    
    if (encryptionInfo.size() >= 4) {
        uint32_t version = encryptionInfo[0] | (encryptionInfo[1]<<8) | (encryptionInfo[2]<<16) | (encryptionInfo[3]<<24);
        if (version == 0x00020003) {
            return Crypto::decryptStandardPackage(encryptionInfo, encryptedPackage, password);
        } else if (version == 0x00040004) {
            return Crypto::decryptAgilePackage(encryptionInfo, encryptedPackage, password);
        }
    }
    
    throw XLInternalError("Unknown encryption version");
}

namespace OpenXLSX::Crypto {

std::vector<uint8_t> readCfbStream(gsl::span<const uint8_t> data, const std::string& streamName) {
    if (data.size() < 512) throw XLInternalError("Invalid CFB size");
    const uint8_t magic[] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
    if (std::memcmp(data.data(), magic, 8) != 0) throw XLInternalError("Not an OLE/CFB file");

    auto getU32 = [](const uint8_t* p) { return p[0] | (p[1]<<8) | (p[2]<<16) | (p[3]<<24); };
    auto getU16 = [](const uint8_t* p) { return p[0] | (p[1]<<8); };
    const uint32_t ENDOFCHAIN = 0xFFFFFFFE;
    const uint32_t FREESECT = 0xFFFFFFFF;

    uint16_t sectorShift = getU16(data.data() + 0x1E);
    uint32_t sectorSize = 1 << sectorShift;
    
    std::vector<uint32_t> difat;
    for(int i=0; i<109; ++i) {
        uint32_t sec = getU32(data.data() + 0x4C + i*4);
        if (sec != FREESECT) difat.push_back(sec);
    }
    
    std::vector<uint32_t> fat;
    for(uint32_t sec : difat) {
        if ((sec + 1) * sectorSize + sectorSize > data.size()) break;
        const uint8_t* p = data.data() + (sec + 1) * sectorSize;
        for(uint32_t i=0; i<sectorSize/4; ++i) fat.push_back(getU32(p + i*4));
    }
    
    auto getChain = [&](uint32_t start) {
        std::vector<uint32_t> chain;
        uint32_t current = start;
        while (current != ENDOFCHAIN && current < fat.size()) {
            chain.push_back(current);
            current = fat[current];
        }
        return chain;
    };

    uint32_t firstMiniFatSector = getU32(data.data() + 0x3C);
    std::vector<uint32_t> minifat;
    if (firstMiniFatSector != ENDOFCHAIN) {
        auto chain = getChain(firstMiniFatSector);
        for(uint32_t sec : chain) {
            if ((sec + 1) * sectorSize + sectorSize > data.size()) break;
            const uint8_t* p = data.data() + (sec + 1) * sectorSize;
            for(uint32_t i=0; i<sectorSize/4; ++i) minifat.push_back(getU32(p + i*4));
        }
    }
    
    uint32_t firstDirSector = getU32(data.data() + 0x30);
    auto dirChain = getChain(firstDirSector);
    uint32_t rootStartSector = 0;
    
    struct Entry { uint32_t start, size; };
    std::map<std::string, Entry> entries;
    
    for(uint32_t sec : dirChain) {
        if ((sec + 1) * sectorSize + sectorSize > data.size()) break;
        const uint8_t* p = data.data() + (sec + 1) * sectorSize;
        for(uint32_t i=0; i<sectorSize/128; ++i) {
            const uint8_t* dir = p + i*128;
            uint16_t nameLen = getU16(dir + 0x40);
            if (nameLen == 0) continue;
            std::string name;
            for(int j=0; j<nameLen-2; j+=2) name += (char)dir[j];
            uint32_t start = getU32(dir + 0x74);
            uint32_t size = getU32(dir + 0x78);
            entries[name] = {start, size};
            if (dir[0x42] == 5) rootStartSector = start;
        }
    }
    
    if (entries.find(streamName) == entries.end()) return std::vector<uint8_t>();
    
    uint32_t miniStreamCutoffSize = getU32(data.data() + 0x38);
    auto rootChain = getChain(rootStartSector);
    auto d = entries[streamName];
    std::vector<uint8_t> out;
    
    if (d.size < miniStreamCutoffSize) {
        auto getMiniChain = [&](uint32_t start) {
            std::vector<uint32_t> chain;
            uint32_t current = start;
            while (current != ENDOFCHAIN && current < minifat.size()) {
                chain.push_back(current);
                current = minifat[current];
            }
            return chain;
        };
        auto chain = getMiniChain(d.start);
        for(uint32_t sec : chain) {
            uint32_t offset = sec * 64;
            if (offset / sectorSize >= rootChain.size()) break;
            uint32_t physSec = rootChain[offset / sectorSize];
            uint32_t physOff = offset % sectorSize;
            if ((physSec + 1) * sectorSize + physOff + 64 > data.size()) break;
            const uint8_t* ptr = data.data() + (physSec + 1) * sectorSize + physOff;
            out.insert(out.end(), ptr, ptr + 64);
        }
    } else {
        auto chain = getChain(d.start);
        for(uint32_t sec : chain) {
            if ((sec + 1) * sectorSize + sectorSize > data.size()) break;
            const uint8_t* ptr = data.data() + (sec + 1) * sectorSize;
            out.insert(out.end(), ptr, ptr + sectorSize);
        }
    }
    
    if (out.size() > d.size) out.resize(d.size);
    return out;
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
    if (encryptionInfo.size() < 24 || encryptedPackage.size() < 8) return std::vector<uint8_t>();

    auto getU32 = [](const uint8_t* p) { return p[0] | (p[1]<<8) | (p[2]<<16) | (p[3]<<24); };

    uint32_t headerSize = getU32(encryptionInfo.data() + 8);
    uint32_t keySize = getU32(encryptionInfo.data() + 28);
    uint32_t saltSize = getU32(encryptionInfo.data() + 12 + headerSize);
    
    if (encryptionInfo.size() < 12 + headerSize + 4 + saltSize) throw XLInternalError("Invalid Standard Encryption Header");
    std::vector<uint8_t> salt(encryptionInfo.data() + 12 + headerSize + 4, encryptionInfo.data() + 12 + headerSize + 4 + saltSize);
    
    std::vector<uint8_t> utf16pw;
    for(char c : password) { utf16pw.push_back(static_cast<uint8_t>(c)); utf16pw.push_back(0); }
    
    std::vector<uint8_t> initialData = salt;
    initialData.insert(initialData.end(), utf16pw.begin(), utf16pw.end());
    
    auto hashSHA1 = [](const std::vector<uint8_t>& d) {
        std::vector<uint8_t> h(SHA_DIGEST_LENGTH);
        SHA1(d.data(), d.size(), h.data());
        return h;
    };
    
    auto H = hashSHA1(initialData);
    
    for(int i = 0; i < 50000; ++i) {
        std::vector<uint8_t> iterData = {static_cast<uint8_t>(i&0xFF), static_cast<uint8_t>((i>>8)&0xFF), static_cast<uint8_t>((i>>16)&0xFF), static_cast<uint8_t>((i>>24)&0xFF)};
        iterData.insert(iterData.end(), H.begin(), H.end());
        H = hashSHA1(iterData);
    }
    
    std::vector<uint8_t> finalData = H;
    finalData.insert(finalData.end(), {0, 0, 0, 0});
    auto H_final = hashSHA1(finalData);
    
    std::vector<uint8_t> buf1(64, 0x36), buf2(64, 0x5C);
    for(int i=0; i<20; ++i) { buf1[i] ^= H_final[i]; buf2[i] ^= H_final[i]; }
    
    auto x1 = hashSHA1(buf1);
    auto x2 = hashSHA1(buf2);
    
    std::vector<uint8_t> keyDerived = x1;
    keyDerived.insert(keyDerived.end(), x2.begin(), x2.end());
    keyDerived.resize(keySize / 8);
    
    OPENSSL_cleanse(utf16pw.data(), utf16pw.size());
    OPENSSL_cleanse(initialData.data(), initialData.size());

    // ECB Decryption
    std::vector<uint8_t> payload(encryptedPackage.begin() + 8, encryptedPackage.end());
    std::vector<uint8_t> decrypted(payload.size());
    
    EVP_CIPHER_CTX *ctx = EVP_CIPHER_CTX_new();
    auto cleanup = gsl::finally([&] { EVP_CIPHER_CTX_free(ctx); });

    const EVP_CIPHER* cipher = nullptr;
    if (keySize == 128) cipher = EVP_aes_128_ecb();
    else if (keySize == 192) cipher = EVP_aes_192_ecb();
    else if (keySize == 256) cipher = EVP_aes_256_ecb();
    else throw XLInternalError("Unsupported AES key size for Standard Encryption");
    
    if (EVP_DecryptInit_ex(ctx, cipher, nullptr, keyDerived.data(), nullptr) != 1) throw XLInternalError("AES init failed");
    EVP_CIPHER_CTX_set_padding(ctx, 0); // ECB uses no padding here
    
    int len1 = 0, len2 = 0;
    EVP_DecryptUpdate(ctx, decrypted.data(), &len1, payload.data(), gsl::narrow_cast<int>(payload.size()));
    EVP_DecryptFinal_ex(ctx, decrypted.data() + len1, &len2);
    
    decrypted.resize(len1 + len2);
    return decrypted;
}

} // namespace OpenXLSX::Crypto
