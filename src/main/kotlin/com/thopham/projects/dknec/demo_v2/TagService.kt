package com.thopham.projects.dknec.demo_v2

import org.springframework.stereotype.Service
import javax.annotation.PostConstruct

@Service
class TagService(val repository: TagRepository) {
    fun findAllByCreatedGreaterThanOrderByCreatedAsc(created: String): List<Tag> {
        return repository.findAllByCreatedGreaterThanOrderByCreated(created)
    }

    fun findAllOrderByCreatedAsc(): List<Tag> {
        return repository.findAllOrderByCreatedAsc()
    }
}