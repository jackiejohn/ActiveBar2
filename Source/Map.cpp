//---------------------------------------------------------------------
//  Copyright (c) 1995-1996, Data Dynamics. All rights reserved.
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//---------------------------------------------------------------------
#include "precomp.h"
#include "assert.h"
#include "map.h"

/*
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
*/

#undef ASSERT
#define ASSERT assert

FMap::FMap(int nBlockSize)
{
assert(nBlockSize > 0);
m_pHashTable = NULL;
m_nHashTableSize = 17;  // default size
m_nCount = 0;
m_pFreeList = NULL;
m_pBlocks = NULL;
m_nBlockSize = nBlockSize;
}

inline UINT FMap::HashKey(void* key) const
{
// default identity hash - works for most primitive values
return ((UINT)(void*)(DWORD)key) >> 4;
}

void FMap::InitHashTable(UINT nHashSize, BOOL bAllocNow)
//
// Used to force allocation of a hash table or to override the default
//   hash table size of (which is fairly small)
{
assert(m_nCount == 0);
assert(nHashSize > 0);

if (m_pHashTable != NULL)
	{
	// free hash table
	delete[] m_pHashTable;
	m_pHashTable = NULL;
	}

if (bAllocNow)
	{
	m_pHashTable = new CAssoc* [nHashSize];
	memset(m_pHashTable, 0, sizeof(CAssoc*) * nHashSize);
	}
m_nHashTableSize = nHashSize;
}

void FMap::RemoveAll()
{
if (m_pHashTable != NULL)
	{
	// free hash table
	delete[] m_pHashTable;
	m_pHashTable = NULL;
	}

m_nCount = 0;
m_pFreeList = NULL;
m_pBlocks->FreeDataChain();
m_pBlocks = NULL;
}

FMap::~FMap()
{
RemoveAll();
}

/////////////////////////////////////////////////////////////////////////////
// Assoc helpers
// same as CList implementation except we store CAssoc's not CNode's
//    and CAssoc's are singly linked all the time

FMap::CAssoc* FMap::NewAssoc()
{
if (m_pFreeList == NULL)
	{
	// add another block
	FPlex* newBlock = FPlex::Create(m_pBlocks, m_nBlockSize, sizeof(FMap::CAssoc));
	// chain them into free list
	FMap::CAssoc* pAssoc = (FMap::CAssoc*) newBlock->data();
	// free in reverse order to make it easier to debug
	pAssoc += m_nBlockSize - 1;
	for (int i = m_nBlockSize-1; i >= 0; i--, pAssoc--)
		{
		pAssoc->pNext = m_pFreeList;
		m_pFreeList = pAssoc;
		}
	}
assert(m_pFreeList != NULL);  // we must have something
FMap::CAssoc* pAssoc = m_pFreeList;
m_pFreeList = m_pFreeList->pNext;
m_nCount++;
assert(m_nCount > 0);  // make sure we don't overflow
memset(&pAssoc->key, 0, sizeof(void*));
memset(&pAssoc->value, 0, sizeof(void*));
return pAssoc;
}

void FMap::FreeAssoc(FMap::CAssoc* pAssoc)
{
pAssoc->pNext = m_pFreeList;
m_pFreeList = pAssoc;
m_nCount--;
assert(m_nCount >= 0);  // make sure we don't underflow
}

FMap::CAssoc* FMap::GetAssocAt(void* key, UINT& nHash) const
// find association (or return NULL)
{
nHash = HashKey(key) % m_nHashTableSize;

if (m_pHashTable == NULL)
	return NULL;

// see if it exists
CAssoc* pAssoc;
for (pAssoc = m_pHashTable[nHash]; pAssoc != NULL; pAssoc = pAssoc->pNext)
	{
	if (pAssoc->key == key)
		return pAssoc;
	}
return NULL;
}

/////////////////////////////////////////////////////////////////////////////

BOOL FMap::Lookup(void* key, void*& rValue) const
{
	UINT nHash;
CAssoc* pAssoc = GetAssocAt(key, nHash);
if (pAssoc == NULL)
	return FALSE;  // not in map

rValue = pAssoc->value;
return TRUE;
}

void*& FMap::operator[](void* key)
{
	UINT nHash;
	CAssoc* pAssoc;
if ((pAssoc = GetAssocAt(key, nHash)) == NULL)
	{
	if (m_pHashTable == NULL)
		InitHashTable(m_nHashTableSize);

	// it doesn't exist, add a new Association
	pAssoc = NewAssoc();
	pAssoc->nHashValue = nHash;
	pAssoc->key = key;
	// 'pAssoc->value' is a constructed object, nothing more

	// put into hash table
	pAssoc->pNext = m_pHashTable[nHash];
	m_pHashTable[nHash] = pAssoc;
	}
return pAssoc->value;  // return new reference
}


BOOL FMap::RemoveKey(void* key)
// remove key - return TRUE if removed
{
if (m_pHashTable == NULL)
	return FALSE;  // nothing in the table

CAssoc** ppAssocPrev;
ppAssocPrev = &m_pHashTable[HashKey(key) % m_nHashTableSize];

CAssoc* pAssoc;
for (pAssoc = *ppAssocPrev; pAssoc != NULL; pAssoc = pAssoc->pNext)
	{
	if (pAssoc->key == key)
		{
		// remove it
		*ppAssocPrev = pAssoc->pNext;  // remove from list
		FreeAssoc(pAssoc);
		return TRUE;
		}
	ppAssocPrev = &pAssoc->pNext;
	}
return FALSE;  // not found
}

/////////////////////////////////////////////////////////////////////////////
// Iterating

void FMap::GetNextAssoc(FPOSITION& rNextPosition,void*& rKey, void*& rValue) const
{
assert(m_pHashTable != NULL);  // never call on empty map
CAssoc* pAssocRet = (CAssoc*)rNextPosition;
assert(pAssocRet != NULL);

if (pAssocRet == (CAssoc*) FSTART_POSITION)
	{
	// find the first association
	for (UINT nBucket = 0; nBucket < m_nHashTableSize; nBucket++)
		if ((pAssocRet = m_pHashTable[nBucket]) != NULL)
			break;
	assert(pAssocRet != NULL);  // must find something
	}

// find next association
	CAssoc* pAssocNext;
if ((pAssocNext = pAssocRet->pNext) == NULL)
	{
	// go to next bucket
	for (UINT nBucket = pAssocRet->nHashValue + 1;
		nBucket < m_nHashTableSize; nBucket++)
			if ((pAssocNext = m_pHashTable[nBucket]) != NULL)
				break;
	}

rNextPosition = (FPOSITION) pAssocNext;

// fill in return data
rKey = pAssocRet->key;
rValue = pAssocRet->value;
}


// FPlex implementation

FPlex* PASCAL FPlex::Create(FPlex*& pHead, UINT nMax, UINT cbElement)
{
ASSERT(nMax > 0 && cbElement > 0);
FPlex* p = (FPlex*) new BYTE[sizeof(FPlex) + nMax * cbElement];
			// may throw exception
p->pNext = pHead;
pHead = p;  // change head (adds in reverse order for simplicity)
return p;
}

void FPlex::FreeDataChain()     // free this one and links
{
FPlex* p = this;
while (p != NULL)
	{
	BYTE* bytes = (BYTE*) p;
	FPlex* pNext = p->pNext;
	delete[] bytes;
	p = pNext;
	}
}
/////////// FPTRLIST

/////////////////////////////////////////////////////////////////////////////

FPtrList::FPtrList(int nBlockSize)
{
	assert(nBlockSize > 0);

	m_nCount = 0;
	m_pNodeHead = m_pNodeTail = m_pNodeFree = NULL;
	m_pBlocks = NULL;
	m_nBlockSize = nBlockSize;
}

void FPtrList::RemoveAll()
{
	// destroy elements
	m_nCount = 0;
	m_pNodeHead = m_pNodeTail = m_pNodeFree = NULL;
	m_pBlocks->FreeDataChain();
	m_pBlocks = NULL;
}

FPtrList::~FPtrList()
{
	RemoveAll();
	assert(m_nCount == 0);
}

/////////////////////////////////////////////////////////////////////////////
// Node helpers
/*
 * Implementation note: CNode's are stored in FPlex blocks and
 *  chained together. Free blocks are maintained in a singly linked list
 *  using the 'pNext' member of CNode with 'm_pNodeFree' as the head.
 *  Used blocks are maintained in a doubly linked list using both 'pNext'
 *  and 'pPrev' as links and 'm_pNodeHead' and 'm_pNodeTail'
 *   as the head/tail.
 *
 * We never free a FPlex block unless the List is destroyed or RemoveAll()
 *  is used - so the total number of FPlex blocks may grow large depending
 *  on the maximum past size of the list.
 */

FPtrList::CNode*
FPtrList::NewNode(FPtrList::CNode* pPrev, FPtrList::CNode* pNext)
{
	if (m_pNodeFree == NULL)
	{
		// add another block
		FPlex* pNewBlock = FPlex::Create(m_pBlocks, m_nBlockSize,
				 sizeof(CNode));

		// chain them into free list
		CNode* pNode = (CNode*) pNewBlock->data();
		// free in reverse order to make it easier to debug
		pNode += m_nBlockSize - 1;
		for (int i = m_nBlockSize-1; i >= 0; i--, pNode--)
		{
			pNode->pNext = m_pNodeFree;
			m_pNodeFree = pNode;
		}
	}
	assert(m_pNodeFree != NULL);  // we must have something

	FPtrList::CNode* pNode = m_pNodeFree;
	m_pNodeFree = m_pNodeFree->pNext;
	pNode->pPrev = pPrev;
	pNode->pNext = pNext;
	m_nCount++;
	assert(m_nCount > 0);  // make sure we don't overflow




	pNode->data = 0; // start with zero

	return pNode;
}

void FPtrList::FreeNode(FPtrList::CNode* pNode)
{

	pNode->pNext = m_pNodeFree;
	m_pNodeFree = pNode;
	m_nCount--;
	assert(m_nCount >= 0);  // make sure we don't underflow

	// if no more elements, cleanup completely
	if (m_nCount == 0)
		RemoveAll();
}

/////////////////////////////////////////////////////////////////////////////

FPOSITION FPtrList::AddHead(void* newElement)
{

	CNode* pNewNode = NewNode(NULL, m_pNodeHead);
	pNewNode->data = newElement;
	if (m_pNodeHead != NULL)
		m_pNodeHead->pPrev = pNewNode;
	else
		m_pNodeTail = pNewNode;
	m_pNodeHead = pNewNode;
	return (FPOSITION) pNewNode;
}

FPOSITION FPtrList::AddTail(void* newElement)
{
	

	CNode* pNewNode = NewNode(m_pNodeTail, NULL);
	pNewNode->data = newElement;
	if (m_pNodeTail != NULL)
		m_pNodeTail->pNext = pNewNode;
	else
		m_pNodeHead = pNewNode;
	m_pNodeTail = pNewNode;
	return (FPOSITION) pNewNode;
}

void FPtrList::AddHead(FPtrList* pNewList)
{
	

	assert(pNewList != NULL);
	
	// add a list of same elements to head (maintain order)
	FPOSITION pos = pNewList->GetTailPosition();
	while (pos != NULL)
		AddHead(pNewList->GetPrev(pos));
}

void FPtrList::AddTail(FPtrList* pNewList)
{
	assert(pNewList != NULL);
	
	// add a list of same elements
	FPOSITION pos = pNewList->GetHeadPosition();
	while (pos != NULL)
		AddTail(pNewList->GetNext(pos));
}

void* FPtrList::RemoveHead()
{
	assert(m_pNodeHead != NULL);  // don't call on empty list !!!
	

	CNode* pOldNode = m_pNodeHead;
	void* returnValue = pOldNode->data;

	m_pNodeHead = pOldNode->pNext;
	if (m_pNodeHead != NULL)
		m_pNodeHead->pPrev = NULL;
	else
		m_pNodeTail = NULL;
	FreeNode(pOldNode);
	return returnValue;
}

void* FPtrList::RemoveTail()
{
	assert(m_pNodeTail != NULL);  // don't call on empty list !!!
	

	CNode* pOldNode = m_pNodeTail;
	void* returnValue = pOldNode->data;

	m_pNodeTail = pOldNode->pPrev;
	if (m_pNodeTail != NULL)
		m_pNodeTail->pNext = NULL;
	else
		m_pNodeHead = NULL;
	FreeNode(pOldNode);
	return returnValue;
}

FPOSITION FPtrList::InsertBefore(FPOSITION position, void* newElement)
{
	

	if (position == NULL)
		return AddHead(newElement); // insert before nothing -> head of the list

	// Insert it before position
	CNode* pOldNode = (CNode*) position;
	CNode* pNewNode = NewNode(pOldNode->pPrev, pOldNode);
	pNewNode->data = newElement;

	if (pOldNode->pPrev != NULL)
	{
		pOldNode->pPrev->pNext = pNewNode;
	}
	else
	{
		assert(pOldNode == m_pNodeHead);
		m_pNodeHead = pNewNode;
	}
	pOldNode->pPrev = pNewNode;
	return (FPOSITION) pNewNode;
}

FPOSITION FPtrList::InsertAfter(FPOSITION position, void* newElement)
{
	
	if (position == NULL)
		return AddTail(newElement); // insert after nothing -> tail of the list

	// Insert it before position
	CNode* pOldNode = (CNode*) position;
	
	CNode* pNewNode = NewNode(pOldNode, pOldNode->pNext);
	pNewNode->data = newElement;

	if (pOldNode->pNext != NULL)
	{
		
		pOldNode->pNext->pPrev = pNewNode;
	}
	else
	{
		assert(pOldNode == m_pNodeTail);
		m_pNodeTail = pNewNode;
	}
	pOldNode->pNext = pNewNode;
	return (FPOSITION) pNewNode;
}

void FPtrList::RemoveAt(FPOSITION position)
{
	
	CNode* pOldNode = (CNode*) position;
	

	// remove pOldNode from list
	if (pOldNode == m_pNodeHead)
	{
		m_pNodeHead = pOldNode->pNext;
	}
	else
	{
		
		pOldNode->pPrev->pNext = pOldNode->pNext;
	}
	if (pOldNode == m_pNodeTail)
	{
		m_pNodeTail = pOldNode->pPrev;
	}
	else
	{
		
		pOldNode->pNext->pPrev = pOldNode->pPrev;
	}
	FreeNode(pOldNode);
}


/////////////////////////////////////////////////////////////////////////////
// slow operations

FPOSITION FPtrList::FindIndex(int nIndex) const
{
	
	assert(nIndex >= 0);

	if (nIndex >= m_nCount)
		return NULL;  // went too far

	CNode* pNode = m_pNodeHead;
	while (nIndex--)
	{
		
		pNode = pNode->pNext;
	}
	return (FPOSITION) pNode;
}

FPOSITION FPtrList::Find(void* searchValue, FPOSITION startAfter) const
{
	

	CNode* pNode = (CNode*) startAfter;
	if (pNode == NULL)
	{
		pNode = m_pNodeHead;  // start at head
	}
	else
	{
		
		pNode = pNode->pNext;  // start after the one specified
	}

	for (; pNode != NULL; pNode = pNode->pNext)
		if (pNode->data == searchValue)
			return (FPOSITION) pNode;
	return NULL;
}

///////////////////////// FARRAY IMPLEMENTATION


FArray::FArray()
{
	m_pData = NULL;
	m_nSize = m_nMaxSize = m_nGrowBy = 0;
}

FArray::~FArray()
{
	if (m_pData!=NULL)
		delete[] (BYTE*)m_pData;
}

HRESULT FArray::SetSize(int nNewSize, int nGrowBy)
{
	assert(nNewSize >= 0);

	if (nGrowBy != -1)
		m_nGrowBy = nGrowBy;  // set new size

	if (nNewSize == 0)
	{
		// shrink to nothing
		if (m_pData!=NULL)
			delete[] (BYTE*)m_pData;
		m_pData = NULL;
		m_nSize = m_nMaxSize = 0;
	}
	else if (m_pData == NULL)
	{
		// create one with exact size
#ifdef SIZE_T_MAX
		assert(nNewSize <= SIZE_T_MAX/sizeof(void*));    // no overflow
#endif
		m_pData = (void**) new BYTE[nNewSize * sizeof(void*)];
		if (m_pData==NULL)
			return E_OUTOFMEMORY;

		memset(m_pData, 0, nNewSize * sizeof(void*));  // zero fill

		m_nSize = m_nMaxSize = nNewSize;
	}
	else if (nNewSize <= m_nMaxSize)
	{
		// it fits
		if (nNewSize > m_nSize)
		{
			// initialize the new elements

			memset(&m_pData[m_nSize], 0, (nNewSize-m_nSize) * sizeof(void*));

		}

		m_nSize = nNewSize;
	}
	else
	{
		// otherwise, grow array
		int nGrowBy = m_nGrowBy;
		if (nGrowBy == 0)
		{
			// heuristically determine growth when nGrowBy == 0
			//  (this avoids heap fragmentation in many situations)
			nGrowBy = min(1024, max(4, m_nSize / 8));
		}
		int nNewMax;
		if (nNewSize < m_nMaxSize + nGrowBy)
			nNewMax = m_nMaxSize + nGrowBy;  // granularity
		else
			nNewMax = nNewSize;  // no slush

		assert(nNewMax >= m_nMaxSize);  // no wrap around
#ifdef SIZE_T_MAX
		assert(nNewMax <= SIZE_T_MAX/sizeof(void*)); // no overflow
#endif
		void** pNewData = (void**) new BYTE[nNewMax * sizeof(void*)];
		if (pNewData==NULL)
			return E_OUTOFMEMORY;

		// copy new data from old
		memcpy(pNewData, m_pData, m_nSize * sizeof(void*));

		// construct remaining elements
		assert(nNewSize > m_nSize);

		memset(&pNewData[m_nSize], 0, (nNewSize-m_nSize) * sizeof(void*));


		// get rid of old stuff (note: no destructors called)
		if (m_pData!=NULL)
			delete[] (BYTE*)m_pData;
		m_pData = pNewData;
		m_nSize = nNewSize;
		m_nMaxSize = nNewMax;
	}
	return S_OK;
}

HRESULT FArray::Append(const FArray& src,int *pRetIndex)
{
	assert(this != &src);   // cannot append to itself
	HRESULT hr;
	int nOldSize = m_nSize;
	hr=SetSize(m_nSize + src.m_nSize);
	if (hr!=S_OK)
		return hr;

	memcpy(m_pData + nOldSize, src.m_pData, src.m_nSize * sizeof(void*));

	if (pRetIndex) 
		*pRetIndex=nOldSize;
	return S_OK;
}

HRESULT FArray::Copy(const FArray& src)
{
	assert(this != &src);   // cannot append to itself
	HRESULT hr;
	hr=SetSize(src.m_nSize);
	if (hr==S_OK)
		memcpy(m_pData, src.m_pData, src.m_nSize * sizeof(void*));
	return hr;
}

void FArray::FreeExtra()
{
	if (m_nSize != m_nMaxSize)
	{
		// shrink to desired size
#ifdef SIZE_T_MAX
		assert(m_nSize <= SIZE_T_MAX/sizeof(void*)); // no overflow
#endif
		void** pNewData = NULL;
		if (m_nSize != 0)
		{
			pNewData = (void**) new BYTE[m_nSize * sizeof(void*)];
			// copy new data from old
			memcpy(pNewData, m_pData, m_nSize * sizeof(void*));
		}

		// get rid of old stuff (note: no destructors called)
		if (m_pData!=NULL)
			delete[] (BYTE*)m_pData;
		m_pData = pNewData;
		m_nMaxSize = m_nSize;
	}
}

/////////////////////////////////////////////////////////////////////////////

HRESULT FArray::SetAtGrow(int nIndex, void* newElement)
{
	assert(nIndex >= 0);
	HRESULT hr;
	if (nIndex >= m_nSize)
	{
		hr=SetSize(nIndex+1);
		if (hr!=S_OK)
			return hr;
	}
	m_pData[nIndex] = newElement;
	return hr;
}

HRESULT FArray::InsertAt(int nIndex, void* newElement, int nCount)
{
	assert(nIndex >= 0);    // will expand to meet need
	assert(nCount > 0);     // zero or negative size not allowed
	HRESULT hr;

	if (nIndex >= m_nSize)
	{
		// adding after the end of the array
		hr=SetSize(nIndex + nCount);  // grow so nIndex is valid
		if (hr!=S_OK)
			return hr;
	}
	else
	{
		// inserting in the middle of the array
		int nOldSize = m_nSize;
		hr=SetSize(m_nSize + nCount);  // grow it to new size
		if (hr!=S_OK)
			return hr;
		// shift old data up to fill gap
		memmove(&m_pData[nIndex+nCount], &m_pData[nIndex],
			(nOldSize-nIndex) * sizeof(void*));

		// re-init slots we copied from

		memset(&m_pData[nIndex], 0, nCount * sizeof(void*));

	}

	// insert new value in the gap
	assert(nIndex + nCount <= m_nSize);
	while (nCount--)
		m_pData[nIndex++] = newElement;
	return S_OK;
}

void FArray::RemoveAt(int nIndex, int nCount)
{
	assert(nIndex >= 0);
	assert(nCount >= 0);
	assert(nIndex + nCount <= m_nSize);

	// just remove a range
	int nMoveCount = m_nSize - (nIndex + nCount);

	if (nMoveCount)
		memcpy(&m_pData[nIndex], &m_pData[nIndex + nCount],
			nMoveCount * sizeof(void*));
	m_nSize -= nCount;
}

HRESULT FArray::InsertAt(int nStartIndex, FArray* pNewArray)
{
	assert(pNewArray != NULL);
	assert(nStartIndex >= 0);
	HRESULT hr;
	if (pNewArray->GetSize() > 0)
	{
		hr=InsertAt(nStartIndex, pNewArray->GetAt(0), pNewArray->GetSize());
		if (hr!=S_OK)
			return hr;
		for (int i = 0; i < pNewArray->GetSize(); i++)
			SetAt(nStartIndex + i, pNewArray->GetAt(i));
	}
	return S_OK;
}

/////////////////////////////////////////////////////////////////////////////
